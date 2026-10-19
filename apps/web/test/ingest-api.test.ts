import { afterEach, expect, test } from "bun:test"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import { deleteMatterSession } from "../src/api/ingest/matters"
import { dispatchMatterContinuation } from "../src/api/ingest/drafting-run"
import { combinedDocumentUrl, fetchCombinedDocument, uploadDocument } from "../src/api/ingest/documents"
import { INGEST_URL } from "../src/config"
import { getDraftingStatus } from "../src/api/ingest/drafting-status"
import type { LeanDraftingStatus } from "../src/api/ingest/lean"

const status = completed as LeanDraftingStatus
const originalFetch = globalThis.fetch

// Bun's fetch type also carries preconnect; keep it so a stub is a whole fetch.
const stubFetch = (respond: () => Promise<Response>): typeof fetch => Object.assign(respond, { preconnect: originalFetch.preconnect })
const originalXhr = globalThis.XMLHttpRequest

afterEach(() => {
  globalThis.fetch = originalFetch
  globalThis.XMLHttpRequest = originalXhr
})

test("slow status polls share one in-flight read per matter and release it on success or failure", async () => {
  const pending: Array<{ resolve: (response: Response) => void; reject: (error: Error) => void }> = []
  globalThis.fetch = stubFetch(() => new Promise<Response>((resolve, reject) => pending.push({ resolve, reject })))
  const first = getDraftingStatus("slow-matter")
  const same = getDraftingStatus("slow-matter")
  const other = getDraftingStatus("other-matter")
  expect(same === first).toBe(true)
  expect(pending).toHaveLength(2)
  pending[0]!.resolve(Response.json(status))
  expect(await first).toEqual(status)
  expect(await same).toEqual(status)
  pending[1]!.resolve(Response.json(status))
  await other
  const failed = getDraftingStatus("slow-matter")
  expect(pending).toHaveLength(3)
  pending[2]!.reject(new Error("Network unavailable"))
  await expect(failed).rejects.toThrow("Network unavailable")
  const retry = getDraftingStatus("slow-matter")
  expect(pending).toHaveLength(4)
  pending[3]!.resolve(Response.json(status))
  expect(await retry).toEqual(status)
})

test("conversation deletion goes through the matter service", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = []
  globalThis.fetch = (async (input, init) => {
    calls.push({ url: String(input), init })
    return Response.json({ ok: true, deleted: true })
  }) as typeof fetch

  await deleteMatterSession("matter-a", "ses_owner/with space")

  expect(calls).toEqual([
    {
      url: expect.stringContaining("/matters/matter-a/sessions/ses_owner%2Fwith%20space"),
      init: { method: "DELETE" },
    },
  ])

  globalThis.fetch = stubFetch(async () => Response.json(
    { error: "This conversation owns an unfinished drafting run and cannot be deleted." },
    { status: 409 },
  ))
  await expect(deleteMatterSession("matter-a", "ses_owner")).rejects.toThrow("unfinished drafting run")
})

test("continuation dispatch goes through the matter backend and validates its claim response", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = []
  globalThis.fetch = (async (input, init) => {
    calls.push({ url: String(input), init })
    return Response.json({
      accepted: true,
      attempted: true,
      existing: false,
      executing: true,
      persisted: true,
      uncertain: false,
    })
  }) as typeof fetch
  const input = { kind: "malformed-tool" as const, progressToken: "malformed:msg_failed", sessionID: "ses_owner" }
  expect(await dispatchMatterContinuation("matter-a", input)).toMatchObject({ accepted: true, persisted: true })
  expect(calls).toHaveLength(1)
  expect(calls[0]).toMatchObject({
    url: expect.stringContaining("/matters/matter-a/continuations/dispatch"),
    init: { method: "POST" },
  })
  expect(JSON.parse(String(calls[0]?.init?.body))).toEqual(input)

  globalThis.fetch = stubFetch(async () => Response.json({ accepted: true }, { status: 200 }))
  await expect(dispatchMatterContinuation("matter-a", input)).rejects.toThrow(
    "Workflow continuation dispatch response is invalid",
  )
})

// Uploads go through XMLHttpRequest for progress events; the stub answers each
// request with one scripted status and body, like the ingest service would.
function scriptXhr(status: number, body: string) {
  globalThis.XMLHttpRequest = class {
    status = 0
    responseText = ""
    timeout = 0
    upload: { onload?: () => void } = {}
    onload?: () => void
    onerror?: () => void
    ontimeout?: () => void
    open() {}
    send() {
      this.status = status
      this.responseText = body
      queueMicrotask(() => this.onload?.())
    }
  } as unknown as typeof XMLHttpRequest
}

test("a failed upload rejects with the ingest service's message instead of resolving its error body", async () => {
  scriptXhr(503, JSON.stringify({ error: "The embedding model could not be loaded.", code: "embedding_model_unavailable" }))
  await expect(uploadDocument("matter", new File(["x"], "source.docx"))).rejects.toThrow(
    "The embedding model could not be loaded.",
  )
  scriptXhr(500, "Internal Server Error")
  await expect(uploadDocument("matter", new File(["x"], "source.docx"))).rejects.toThrow("Could not read source.docx (500)")
})

test("a combined document downloads from ingest's combined route, and its refusals read as plain reasons", async () => {
  expect(combinedDocumentUrl("matter-a")).toBe(`${INGEST_URL}/matters/matter-a/documents/combined`)
  const calls: string[] = []
  globalThis.fetch = (async (input) => {
    calls.push(String(input))
    return new Response(new Uint8Array([80, 75, 3, 4]), { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } })
  }) as typeof fetch
  const file = await fetchCombinedDocument("matter-a", "Supply note.docx")
  expect(calls).toEqual([`${INGEST_URL}/matters/matter-a/documents/combined`])
  expect([...new Uint8Array(await file.arrayBuffer())]).toEqual([80, 75, 3, 4])

  const refuse = (status: number, body: unknown) => {
    globalThis.fetch = stubFetch(async () => Response.json(body, { status }))
    return fetchCombinedDocument("matter-a", "Supply note.docx")
  }
  await expect(refuse(409, { error: "combined_parts_missing" })).rejects.toThrow(
    "Could not download Supply note.docx: the draft or the memo it is built from is no longer in the matter.",
  )
  await expect(refuse(409, { error: "unsupported_memo_dependency", detail: "unsupported_memo_numbering" })).rejects.toThrow(
    "Could not download Supply note.docx: the memo holds content the combined file cannot carry. Download the draft and the memo separately.",
  )
  await expect(refuse(404, { error: "no_combined_output" })).rejects.toThrow("Could not download Supply note.docx: this drafting run has no combined file.")
  // Any other refusal keeps the route's own message.
  await expect(refuse(409, { error: "The matter is busy." })).rejects.toThrow("The matter is busy.")
  globalThis.fetch = stubFetch(async () => new Response("Internal Server Error", { status: 500 }))
  await expect(fetchCombinedDocument("matter-a", "Supply note.docx")).rejects.toThrow("Could not download Supply note.docx (500)")
})

test("each combined refusal reads as a plain reason, and an unreadable Word file names the draft or the memo", async () => {
  const refuse = (status: number, body: unknown) => {
    globalThis.fetch = stubFetch(async () => Response.json(body, { status }))
    return fetchCombinedDocument("matter-a", "Supply note.docx").then(
      () => "downloaded",
      (error: Error) => error.message,
    )
  }
  const prefix = "Could not download Supply note.docx: "
  const message = "internal composer message"

  expect(await refuse(409, { error: "empty_memo", message })).toBe(`${prefix}the memo has no text to place before the draft.`)
  expect(await refuse(409, { error: "unsupported_memo_numbering", message })).toBe(
    `${prefix}the memo uses list numbering the combined file cannot carry. Download the draft and the memo separately.`,
  )
  expect(await refuse(409, { error: "unsupported_memo_namespace", message })).toBe(
    `${prefix}the memo uses Word formatting the combined file cannot carry. Download the draft and the memo separately.`,
  )
  // A memo refusal the map does not name yet still reads plainly.
  expect(await refuse(409, { error: "unsupported_memo_footnotes", message })).toBe(
    `${prefix}the memo holds content the combined file cannot carry. Download the draft and the memo separately.`,
  )
  for (const error of ["invalid_docx_package", "invalid_docx_xml"]) {
    expect(await refuse(409, { error, part: "body", message })).toBe(
      `${prefix}the draft could not be read as a Word document. Open it in Word, save it again, and try again.`,
    )
    expect(await refuse(409, { error, part: "memo", message })).toBe(
      `${prefix}the memo could not be read as a Word document. Open it in Word, save it again, and try again.`,
    )
    // A refusal without a known part still reads plainly.
    expect(await refuse(409, { error })).toBe(
      `${prefix}the draft or the memo could not be read as a Word document. Open it in Word, save it again, and try again.`,
    )
  }
  // Busy is reported once; the client does not retry on its own.
  let calls = 0
  globalThis.fetch = stubFetch(async () => {
    calls++
    return Response.json({ error: "combined_parts_busy", message }, { status: 503, headers: { "Retry-After": "5" } })
  })
  await expect(fetchCombinedDocument("matter-a", "Supply note.docx")).rejects.toThrow(
    `${prefix}the draft or the memo is being updated. Try again shortly.`,
  )
  expect(calls).toBe(1)
})
