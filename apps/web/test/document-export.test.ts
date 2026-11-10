import { afterEach, expect, test } from "bun:test"
import { documentDownloadUrl, fetchDocumentExportBytes } from "../src/api/ingest/exports"
import { fetchVersionBytes, VersionUnavailableError } from "../src/api/ingest/versions"
import { fetchBlacklineBytes } from "../src/api/ingest/compare"
import { exportChronology } from "../src/api/ingest/chronology"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import { isLeanDraftingStatus } from "../src/api/ingest/lean"
import { draftedGroups } from "../src/components/drafting/model"
import DraftedList from "../src/components/drafting/DraftedList"
import ResultCard from "../src/components/narration/ResultCard"

const originalFetch = globalThis.fetch
afterEach(() => { globalThis.fetch = originalFetch })

test("ordinary sharing copies use explicit encoded export requests while archival source reads remain raw", async () => {
  const requests: URL[] = [], bytes = new Uint8Array([80, 75, 3, 4])
  globalThis.fetch = (async input => { requests.push(new URL(String(input))); return new Response(bytes) }) as typeof fetch
  expect(await fetchDocumentExportBytes("matter/one", { name: "A & B.docx" })).toEqual(bytes)
  await fetchDocumentExportBytes("matter/one", { versionId: "old/id" })
  await fetchVersionBytes("matter/one", "old/id")
  expect(requests[0]!.pathname).toEndWith("/matters/matter%2Fone/documents/export")
  expect(requests[0]!.searchParams.get("name")).toBe("A & B.docx")
  expect(requests[1]!.searchParams.get("versionId")).toBe("old/id")
  expect(requests[1]!.searchParams.has("name")).toBe(false)
  expect(requests[2]!.pathname).toEndWith("/documents/versions/old%2Fid/content")
})

test("proposal downloads retain the renderer disclosure and retained export refusals keep their source status", async () => {
  let notice: string | undefined, request: URL | undefined
  globalThis.fetch = (async input => {
    request = new URL(String(input))
    return new Response(new Uint8Array([80, 75]), { headers: { "X-DocHaus-Revision-Notice": encodeURIComponent("Original revisions: check the source.") } })
  }) as typeof fetch
  await fetchDocumentExportBytes("one", { name: "Agreement.docx", proposals: true }, value => { notice = value })
  expect(request!.searchParams.get("proposals")).toBe("1")
  expect(notice).toBe("Original revisions: check the source.")
  globalThis.fetch = (async () => Response.json({ error: "Removed by redaction.", code: "removed_by_redaction" }, { status: 410 })) as unknown as typeof fetch
  const error = await fetchDocumentExportBytes("one", { versionId: "old" }).catch(error => error)
  expect(error).toBeInstanceOf(VersionUnavailableError)
  expect(error).toMatchObject({ code: "removed_by_redaction", status: 410 })
})

test("blackline export preparation is requested before display/download without changing raw comparison callers", async () => {
  const requests: URL[] = []
  globalThis.fetch = (async input => { requests.push(new URL(String(input))); return new Response(new Uint8Array([80, 75])) }) as typeof fetch
  const params = { left: "Before.docx", right: "After.docx", leftVersion: "base", rightVersion: "source" }
  await fetchBlacklineBytes("one", params, { exportCopy: true })
  await fetchBlacklineBytes("one", params)
  expect(requests[0]!.searchParams.get("exportCopy")).toBe("1")
  expect(requests[0]!.searchParams.get("rightVersion")).toBe("source")
  expect(requests[1]!.searchParams.has("exportCopy")).toBe(false)
})

test("chronology browser download prepares its exact newly registered version rather than re-reading latest by name", async () => {
  const requests: { url: URL; method: string }[] = [], raw = new Uint8Array([1, 2]), prepared = new Uint8Array([3, 4])
  globalThis.fetch = (async (input, init) => {
    const url = new URL(String(input)); requests.push({ url, method: init?.method ?? "GET" })
    if (requests.length === 1) return new Response(raw, { headers: { "X-DocHaus-Export-Name": "Chronology-snapshot.docx", "X-DocHaus-Export-Version": "actual-created-version" } })
    return new Response(prepared)
  }) as typeof fetch
  const result = await exportChronology("matter/one", { expectedRevision: 7, scope: "all-visible", includeHidden: false })
  expect(result.name).toBe("Chronology-snapshot.docx")
  expect(new Uint8Array(await result.blob.arrayBuffer())).toEqual(prepared)
  expect(requests.map(request => request.method)).toEqual(["POST", "GET"])
  expect(requests[1]!.url.pathname).toEndWith("/documents/export")
  expect(requests[1]!.url.searchParams.get("versionId")).toBe("actual-created-version")
  expect(requests[1]!.url.searchParams.has("name")).toBe(false)
})

test("drafting list and completion-card Word download anchors use sharing copies while other formats keep raw downloads", () => {
  if (!isLeanDraftingStatus(completed)) throw new Error("Expected a valid reviewed synthetic completed status")
  const views = [
    renderToStaticMarkup(createElement(DraftedList, { matterId: "one", groups: draftedGroups(completed), onView: () => {} })),
    renderToStaticMarkup(createElement(ResultCard, { status: completed, matterId: "one", onViewDocument: () => {} })),
  ]
  for (const html of views) {
    const anchors = [...html.matchAll(/<a [^>]*href="([^"]*)"[^>]*download="([^"]*\.docx)"/g)]
    expect(anchors.length).toBeGreaterThan(0)
    for (const [, href, name] of anchors) {
      const url = new URL(href!.replaceAll("&amp;", "&"))
      expect(url.pathname).toEndWith("/documents/export")
      expect(url.searchParams.get("name")).toBe(name)
    }
  }
  expect(new URL(documentDownloadUrl("one", "Evidence.pdf")).pathname).toEndWith("/documents/content")
})
