import { afterEach, expect, test } from "bun:test"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import { createMatter, deleteMatterSession, renameMatter } from "../src/api/ingest/matters"
import { dispatchMatterContinuation } from "../src/api/ingest/drafting-run"
import { DocumentExistsError, DuplicateUploadError, uploadDocument } from "../src/api/ingest/documents"
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

test("matter saves send the client only when the caller passes one", async () => {
  const bodies: unknown[] = []
  globalThis.fetch = (async (_input, init) => {
    bodies.push(JSON.parse(String(init?.body)))
    return Response.json({ id: "matter-a", title: "Supply", dir: "/m", created_at: 1 })
  }) as typeof fetch

  await renameMatter("matter-a", "Supply", "2026-0042", ["EW"], "playbook-nda")
  await renameMatter("matter-a", "Supply", "2026-0042", ["EW"], "playbook-nda", "Northwind Traders Ltd")
  await renameMatter("matter-a", "Supply", "2026-0042", ["EW"], "playbook-nda", "")
  await createMatter("Supply", undefined, undefined, undefined, "Northwind Traders Ltd")

  expect(bodies[0]).not.toHaveProperty("client")
  expect(bodies[1]).toMatchObject({ client: "Northwind Traders Ltd" })
  expect(bodies[2]).toMatchObject({ client: "" })
  expect(bodies[3]).toMatchObject({ title: "Supply", client: "Northwind Traders Ltd" })
})

test("matter saves send the posture only when the caller passes one", async () => {
  const bodies: unknown[] = []
  globalThis.fetch = (async (_input, init) => {
    bodies.push(JSON.parse(String(init?.body)))
    return Response.json({ id: "matter-a", title: "Supply", dir: "/m", created_at: 1 })
  }) as typeof fetch

  await renameMatter("matter-a", "Supply", "2026-0042", ["EW"], "playbook-nda", "Northwind Traders Ltd")
  await renameMatter("matter-a", "Supply", "2026-0042", ["EW"], "playbook-nda", undefined, "conservative")
  await renameMatter("matter-a", "Supply", "2026-0042", ["EW"], "playbook-nda", undefined, "")
  await createMatter("Supply", undefined, undefined, undefined, "Northwind Traders Ltd", "client-favorable")
  await createMatter("Supply")

  expect(bodies[0]).not.toHaveProperty("posture")
  expect(bodies[1]).toMatchObject({ posture: "conservative" })
  expect(bodies[1]).not.toHaveProperty("client")
  expect(bodies[2]).toMatchObject({ posture: "" })
  expect(bodies[3]).toMatchObject({ title: "Supply", client: "Northwind Traders Ltd", posture: "client-favorable" })
  expect(bodies[4]).not.toHaveProperty("posture")
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
// request with one scripted status and body, like the ingest service would, and
// keeps the forms it was sent.
function scriptXhr(status: number, body: string) {
  const sent: FormData[] = []
  globalThis.XMLHttpRequest = class {
    status = 0
    responseText = ""
    timeout = 0
    upload: { onload?: () => void } = {}
    onload?: () => void
    onerror?: () => void
    ontimeout?: () => void
    open() {}
    send(form: FormData) {
      sent.push(form)
      this.status = status
      this.responseText = body
      queueMicrotask(() => this.onload?.())
    }
  } as unknown as typeof XMLHttpRequest
  return sent
}

test("a failed upload rejects with the ingest service's message instead of resolving its error body", async () => {
  scriptXhr(503, JSON.stringify({ error: "The embedding model could not be loaded.", code: "embedding_model_unavailable" }))
  await expect(uploadDocument("matter", new File(["x"], "source.docx"))).rejects.toThrow(
    "The embedding model could not be loaded.",
  )
  scriptXhr(500, "Internal Server Error")
  await expect(uploadDocument("matter", new File(["x"], "source.docx"))).rejects.toThrow("Could not read source.docx (500)")
})

test("a same-name upload is refused with the existing name until the caller chooses replace or keep both", async () => {
  const file = new File(["theirs"], "agreement.docx")
  const refused = scriptXhr(409, JSON.stringify({
    error: "Agreement.docx is already in this matter. Replace it, or keep both and add the upload under a new name.",
    code: "source_exists",
    name: "Agreement.docx",
  }))
  const error = await uploadDocument("matter", file).catch((error: unknown) => error)
  expect(error).toBeInstanceOf(DocumentExistsError)
  expect(error).toMatchObject({ existing: "Agreement.docx", role: "source" })
  expect((error as DocumentExistsError).message).toStartWith("Agreement.docx is already in this matter.")
  expect((refused[0]?.get("file") as File).name).toBe("agreement.docx")
  expect(refused[0]?.has("existing")).toBe(false)

  // Keep both resolves with the name the upload was stored under.
  const kept = scriptXhr(200, JSON.stringify({ name: "Agreement (2).docx", docPath: "/matter/Agreement (2).docx", sections: 1, chunks: 1, injection: null }))
  expect((await uploadDocument("matter", file, { existing: "keep-both" })).name).toBe("Agreement (2).docx")
  expect(kept[0]?.get("existing")).toBe("keep-both")
  const replaced = scriptXhr(200, JSON.stringify({ name: "Agreement.docx", docPath: "/matter/Agreement.docx", sections: 1, chunks: 1, injection: null }))
  expect((await uploadDocument("matter", file, { existing: "replace" })).name).toBe("Agreement.docx")
  expect(replaced[0]?.get("existing")).toBe("replace")

  // A generated deliverable of that name is named too, so the caller can keep both.
  scriptXhr(409, JSON.stringify({
    error: "Cannot replace generated deliverable Draft.docx with a source document.",
    code: "deliverable_exists",
    name: "Draft.docx",
  }))
  const deliverable = await uploadDocument("matter", new File(["x"], "Draft.docx")).catch((error: unknown) => error)
  expect(deliverable).toBeInstanceOf(DocumentExistsError)
  expect(deliverable).toMatchObject({ existing: "Draft.docx", role: "deliverable" })

  // Any other conflict stays a plain failure.
  scriptXhr(409, JSON.stringify({ error: "Alias.docx is a symbolic-link alias, not a matter document." }))
  const alias = await uploadDocument("matter", new File(["x"], "Alias.docx")).catch((error: unknown) => error)
  expect(alias).not.toBeInstanceOf(DocumentExistsError)
  expect((alias as Error).message).toBe("Alias.docx is a symbolic-link alias, not a matter document.")
})

// #268: bytes the same-name document already keeps are refused as a duplicate,
// whatever the choice, naming the version that holds them.
test("an identical upload is refused as a duplicate of the version that holds its bytes", async () => {
  scriptXhr(409, JSON.stringify({
    error: "Agreement.docx is already in this matter as version 2 of Agreement.docx. Nothing new was stored.",
    code: "duplicate", name: "Agreement.docx", document_id: "doc-1", version: { id: "v-2", number: 2, current: false },
  }))
  const error = await uploadDocument("matter", new File(["same"], "Agreement.docx"), { existing: "replace" }).catch((error: unknown) => error)
  expect(error).toBeInstanceOf(DuplicateUploadError)
  expect(error).toMatchObject({ existing: "Agreement.docx", version: { id: "v-2", number: 2, current: false } })
  // A file from before versions were kept has no version to name.
  scriptXhr(409, JSON.stringify({ error: "Old.docx is already in this matter as Old.docx. Nothing new was stored.", code: "duplicate", name: "Old.docx", document_id: null, version: null }))
  expect(await uploadDocument("matter", new File(["same"], "Old.docx")).catch((error: unknown) => error)).toMatchObject({ existing: "Old.docx", version: null })
})
