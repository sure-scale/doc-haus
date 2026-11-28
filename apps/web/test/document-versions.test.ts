import { afterEach, describe, expect, test } from "bun:test"
import { VERSION_ORIGINS } from "../../../services/ingest/src/document-versions"
import { compareUrl } from "../src/api/ingest/compare"
import { fetchVersionBytes, restoreVersion, RestoreRefusedError, VersionUnavailableError, type DocumentVersion } from "../src/api/ingest/versions"
import { choiceParams } from "../src/components/CompareDialog"
import { branchNote, previousVersion, unavailableReason, versionActor, versionOption, versionOrigin } from "../src/components/DocumentVersions"

const originalFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = originalFetch
})

const at = Date.UTC(2026, 9, 24, 9)
const version = (number: number, patch: Partial<DocumentVersion> = {}): DocumentVersion => ({
  id: `v-${number}`, number, sha256: String(number).repeat(64), ext: ".docx", bytes: 1000, parent: number > 1 ? `v-${number - 1}` : null,
  derived_from: null, origin: "revision", actor: "unknown", run: null, label: null, created_at: at + number * 60_000, current: false, ...patch,
})

describe("version metadata helpers", () => {
  test("every ingest origin has a specific description while unknown future origins fall back cleanly", () => {
    const words = (origin: string, parent: string | null = "v-0") => versionOrigin({ id: "v", origin: origin as DocumentVersion["origin"], parent })
    expect(words("from-a-newer-ingest")).toBe("Saved")
    for (const origin of VERSION_ORIGINS) {
      for (const parent of [null, "v-0"]) expect(words(origin, parent)).not.toBe("Saved")
    }
    expect(words("system", null)).not.toBe(words("system", "v-1"))
    expect(words("unknown", null)).not.toBe(words("unknown", "v-1"))
  })

  test("picker options omit the year that the full title retains", () => {
    const option = versionOption(version(3, { origin: "revision", created_at: Date.UTC(2026, 9, 24, 9) }), [])
    expect(option.text).not.toContain("2026")
    expect(option.title).toContain("2026")
  })

  test("restores identify the restored version number when the restore event supplies it", () => {
    const restored = version(5, { origin: "restore" })
    const event = { id: 1, kind: "restore" as const, version_id: "v-5", actor: "unknown", detail: { restored_from: "v-2", restored_number: 2 }, created_at: at }
    expect(versionOrigin(restored, [event])).toContain("version 2")
    expect(versionOrigin(restored, [])).not.toContain("version 2")
  })

  test("pure reply-only labels override the default revision origin while mixed labels do not", () => {
    expect(versionOrigin(version(2, { label: "1 reply added" }))).toBe("1 reply added")
    expect(versionOrigin(version(3, { label: "2 replies added" }))).toBe("2 replies added")
    expect(versionOrigin(version(5, { label: "Signed copy, 2 replies added" }))).toBe(versionOrigin(version(4)))
  })

  test("only named human actors are returned", () => {
    expect(["unknown", "", "doc.haus", "agent:legal-review", null].map(versionActor)).toEqual([undefined, undefined, undefined, undefined, undefined])
    expect(versionActor(" Pat Reviewer ")).toBe("Pat Reviewer")
  })

  test("siblings branching from the same parent both record their common ancestor", () => {
    const list = [version(4, { parent: "v-2" }), version(3), version(2), version(1)]
    expect(list.map((v) => branchNote(v, list))).toEqual(["Based on version 2", "Based on version 2", undefined, undefined])
  })

  test("previousVersion compares against the prior version or the current version when edited externally", () => {
    const list = [version(3, { current: true }), version(2), version(1)]
    expect(previousVersion({ versions: list, working_version: "v-3" })?.id).toBe("v-2")
    expect(previousVersion({ versions: list, working_version: null })?.id).toBe("v-3")
    expect(previousVersion({ versions: [version(1, { current: true })], working_version: "v-1" })).toBeUndefined()
    expect(previousVersion({ versions: [], working_version: null })).toBeUndefined()
  })
})

describe("version requests", () => {
  test("compare sends a kept version per side, after the names", () => {
    expect(new URL(compareUrl("m-1", { left: "A.docx", right: "A.docx", leftVersion: "v-2", rightVersion: "v-3" })).search).toBe(
      "?left=A.docx&right=A.docx&leftVersion=v-2&rightVersion=v-3&ignoreFormatting=1&moves=1&headers=1&format=docx",
    )
    const options = { ignoreFormatting: true, detectMoves: true, headersFooters: true }
    expect(choiceParams({ name: "A.docx", other: "A.docx", otherVersion: "v-2", thisSide: "right", thisLabel: "A (version 3)", otherLabel: "A (version 2)", options })).toEqual({
      left: "A.docx", right: "A.docx", rightLabel: "A (version 3)", leftVersion: "v-2", leftLabel: "A (version 2)", ...options,
    })
    expect(choiceParams({ name: "A.docx", other: "B.docx", thisVersion: "v-1", thisSide: "left", thisLabel: "", otherLabel: "", options })).toEqual({
      left: "A.docx", right: "B.docx", leftVersion: "v-1", ...options,
    })
  })

  test("an unreadable or removed version's copy is distinguished from generic server failures", async () => {
    globalThis.fetch = (async () => Response.json({ error: "The kept bytes of this version are missing or damaged.", code: "version_damaged" }, { status: 409 })) as unknown as typeof fetch
    const damaged = await fetchVersionBytes("m-1", "v-2").catch((error: unknown) => error)
    expect(damaged).toBeInstanceOf(VersionUnavailableError)
    expect(unavailableReason(damaged)).toBeDefined()

    globalThis.fetch = (async () => Response.json({ error: "Removed by redaction.", code: "removed_by_redaction" }, { status: 410 })) as unknown as typeof fetch
    const redacted = await fetchVersionBytes("m-1", "v-2").catch((error: unknown) => error)
    expect(redacted).toMatchObject({ status: 410, code: "removed_by_redaction" })
    expect(unavailableReason(redacted)).toContain("redacted")

    globalThis.fetch = (async () => new Response("gone", { status: 404 })) as unknown as typeof fetch
    expect(unavailableReason(await fetchVersionBytes("m-1", "v-2").catch((error: unknown) => error))).toBeDefined()

    globalThis.fetch = (async () => new Response("down", { status: 502 })) as unknown as typeof fetch
    expect(await fetchVersionBytes("m-1", "v-2").catch((error: unknown) => error)).not.toBeInstanceOf(VersionUnavailableError)
  })

  test("restore posts to the version's route and a refusal keeps the route's words and code", async () => {
    const calls: { url: string; method?: string }[] = []
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      calls.push({ url: String(input), method: init?.method })
      return Response.json({ name: "A.docx", version: { id: "v-5", number: 5 }, restored_from: { id: "v-2", number: 2 } })
    }) as typeof fetch
    expect((await restoreVersion("m 1", "v-2")).version?.number).toBe(5)
    expect(calls).toHaveLength(1)
    expect(calls[0]!.url).toEndWith("/matters/m%201/documents/versions/v-2/restore")
    expect(calls[0]!.method).toBe("POST")

    globalThis.fetch = (async () => Response.json({ error: "A.docx already holds version 2.", code: "already_current" }, { status: 409 })) as unknown as typeof fetch
    const refused = await restoreVersion("m-1", "v-2").catch((error: unknown) => error)
    expect(refused).toBeInstanceOf(RestoreRefusedError)
    expect(refused).toMatchObject({ code: "already_current", message: "A.docx already holds version 2." })
    expect(unavailableReason(refused)).toBeUndefined()

    globalThis.fetch = (async () => Response.json({ error: "Removed by redaction.", code: "removed_by_redaction" }, { status: 410 })) as unknown as typeof fetch
    expect(unavailableReason(await restoreVersion("m-1", "v-2").catch((error: unknown) => error))).toContain("redacted")
  })
})
