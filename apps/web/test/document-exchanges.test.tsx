import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, type ReactElement } from "react"
import type { DocumentVersions, DocumentVersion, LinkedDocumentBranch } from "../src/api/ingest/versions"
import type { DocumentExchange, ReturnCandidate, ReturnInspection } from "../src/api/ingest/exchanges"
import type { Document } from "../src/api/ingest/documents"

const version = (id: string, number: number, patch: Partial<DocumentVersion> = {}): DocumentVersion => ({ id, number, sha256: String(number).repeat(64), ext: ".docx", bytes: 100, parent: number === 1 ? null : `v${number - 1}`, derived_from: null, origin: "upload", actor: "Avery", run: null, label: null, created_at: number, current: number === 3, ...patch })
const versions = [version("v3", 3), version("v2", 2), version("v1", 1)]
const saved = (name = "Agreement.docx", items = versions, working: string | null = "v3"): DocumentVersions => ({ name, document: { id: `doc:${name}`, name, current_version: working }, versions: items, events: [], working_version: working })
const candidate: ReturnCandidate = { documentId: "doc:Agreement.docx", documentName: "Agreement.docx", versionId: "v1", versionNumber: 1, sha256: versions[2]!.sha256, current: false, sent: true, signals: ["matching-id"], score: 1, reason: "Matches the sent version identity." }
const inspection = (patch: Partial<ReturnInspection> = {}): ReturnInspection => ({ kind: "matched", sha256: "f".repeat(64), candidates: [candidate], selected: candidate, reason: "This return matches an earlier sent copy.", warnings: [], complete: true, ...patch })
type Request = { url: URL; method: string; init?: RequestInit }

async function render(element: () => Promise<ReactElement>, respond: (request: Request) => Promise<Response> | Response) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Element", "Node", "Event", "HTMLInputElement", "HTMLSelectElement", "HTMLDialogElement", "getComputedStyle", "IS_REACT_ACT_ENVIRONMENT"] as const
  const descriptors = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const fetchBefore = globalThis.fetch, requests: Request[] = [], downloads: string[] = []
  browser.HTMLAnchorElement.prototype.click = function () { downloads.push(this.download) }
  globalThis.fetch = (async (input, init) => {
    const request = { url: new URL(String(input), browser.location.href), method: init?.method ?? "GET", init }; requests.push(request)
    return respond(request)
  }) as typeof fetch
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div"); document.body.append(container)
  const root = createRoot(container)
  await act(async () => { root.render(await element()) })
  const button = (text: string) => [...container.querySelectorAll<HTMLButtonElement>("button")].find(item => item.textContent === text)
  const field = <T extends HTMLElement>(label: string) => [...container.querySelectorAll("label")].find(item => item.textContent?.startsWith(label))?.querySelector<T>("input,select")
  const click = async (text: string) => { const target = button(text); expect(target).toBeDefined(); await act(async () => { target!.click() }) }
  const select = async (label: string, value: string) => {
    const target = field<HTMLSelectElement>(label); expect(target).toBeDefined()
    await act(async () => { target!.value = value; target!.dispatchEvent(new Event("change", { bubbles: true })) })
  }
  const upload = async (file = new File(["reviewed synthetic returned content"], "Renamed return.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })) => {
    const target = container.querySelector<HTMLInputElement>('input[type="file"]')!
    await act(async () => { Object.defineProperty(target, "files", { configurable: true, value: [file] }); target.dispatchEvent(new Event("change", { bubbles: true })) })
    return file
  }
  const input = async (label: string, value: string) => {
    const target = field<HTMLInputElement>(label); expect(target).toBeDefined()
    await act(async () => { Object.getOwnPropertyDescriptor(browser.HTMLInputElement.prototype, "value")!.set!.call(target, value); target!.dispatchEvent(new Event("input", { bubbles: true })) })
  }
  return { browser, root, container, requests, downloads, field, button, click, select, upload, input,
    text: () => container.textContent ?? "",
    async update(element: ReactElement) { await act(async () => { root.render(element) }) },
    async cleanup() { await act(async () => root.unmount()); globalThis.fetch = fetchBefore; await browser.happyDOM.close(); for (const [name, descriptor] of descriptors) { if (descriptor) Object.defineProperty(globalThis, name, descriptor); else delete (globalThis as Record<string, unknown>)[name] } },
  }
}

test("send starts from the current saved copy, obeys matching defaults and sends the explicitly chosen blackline once", async () => {
  let refreshed = 0
  const view = await render(async () => {
    const { default: DocumentExchanges } = await import("../src/components/DocumentExchanges")
    return createElement(DocumentExchanges, { matterId: "matter/one", name: "Agreement.docx", versions, refresh: () => { refreshed++ } })
  }, request => {
    if (request.url.pathname.endsWith("/preferences")) return Response.json({ documentMatchingId: false })
    if (request.url.pathname.endsWith("/exchanges")) return Response.json({ exchanges: [] })
    if (request.url.pathname.endsWith("/v2/send")) return new Response(new Uint8Array([1,2,3]), { headers: { "X-DocHaus-Draft-Number": "1" } })
    throw new Error(`Unexpected UI request: ${request.method} ${request.url.pathname}`)
  })
  try {
    expect(view.requests.filter(request => request.method === "POST")).toHaveLength(0)
    await view.click("Record a send")
    expect(view.field<HTMLSelectElement>("Version")!.value).toBe("v3")
    expect(view.field<HTMLInputElement>("Include an opaque matching ID")!.checked).toBe(false)
    expect(view.text()).toContain("does not email or share it")
    expect(view.text()).toContain("Unsaved changes in Word are not included")
    await view.select("Version", "v2"); await view.select("Comparison copy", "v1")
    await view.click("Record as sent and download")
    const writes = view.requests.filter(request => request.method === "POST")
    expect(writes).toHaveLength(1); expect(writes[0]!.url.pathname).toContain("/matters/matter%2Fone/documents/versions/v2/send")
    expect(JSON.parse(writes[0]!.init!.body as string)).toMatchObject({ marker: false, baseVersionId: "v1" })
    expect(typeof JSON.parse(writes[0]!.init!.body as string).requestId).toBe("string")
    expect(view.downloads).toEqual(["Agreement.docx"]); expect(refreshed).toBe(1)
    expect(Boolean(view.container.querySelector("dialog"))).toBe(false)
  } finally { await view.cleanup() }
})

test("a send waits for its matching preference and a failed acknowledgement retries the identical request identity", async () => {
  let preference!: (response: Response) => void, attempts = 0
  const view = await render(async () => {
    const { default: DocumentExchanges } = await import("../src/components/DocumentExchanges")
    return createElement(DocumentExchanges, { matterId: "matter", name: "Agreement.docx", versions })
  }, request => {
    if (request.url.pathname.endsWith("/preferences")) return new Promise(resolve => { preference = resolve })
    if (request.url.pathname.endsWith("/exchanges")) return Response.json({ exchanges: [] })
    if (request.url.pathname.endsWith("/send")) { attempts++; return attempts === 1 ? Response.json({ error: "Acknowledgement unavailable. Try again." }, { status: 503 }) : new Response("exact copy") }
    throw new Error(`Unexpected UI request: ${request.method} ${request.url.pathname}`)
  })
  try {
    await view.click("Record a send")
    expect(view.button("Record as sent and download")?.disabled).toBe(true)
    await act(async () => { preference(Response.json({ documentMatchingId: true })) })
    expect(view.field<HTMLInputElement>("Include an opaque matching ID")!.checked).toBe(true)
    await view.click("Record as sent and download")
    expect(view.text()).toContain("Acknowledgement unavailable")
    await view.click("Record as sent and download")
    const writes = view.requests.filter(request => request.method === "POST")
    expect(writes).toHaveLength(2); expect(writes[0]!.init!.body).toBe(writes[1]!.init!.body)
  } finally { await view.cleanup() }
})

test("a return matched to an older sent round preserves that base while comparing the new saved version", async () => {
  let refreshed = 0
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }], onNew: () => { throw new Error("not a new document") }, onSaved: () => { refreshed++ } })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection())
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved())
    if (request.url.pathname.endsWith("/returns")) return Response.json({ name: "Agreement.docx", version: { id: "returned-v4", number: 4 }, comparison: { leftVersion: "v1", rightVersion: "returned-v4" } })
    if (request.url.pathname.endsWith("/compare")) return Response.json({ message: "Synthetic UI fixture does not render Word bytes." }, { status: 409 })
    throw new Error(`Unexpected UI request: ${request.method} ${request.url.pathname}`)
  })
  try {
    await view.upload()
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("v1")
    expect(view.text()).toContain("Its parent stays version 1; existing internal versions are kept")
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns"))).toHaveLength(0)
    await view.click("Save returned version and compare")
    const writes = view.requests.filter(request => request.url.pathname.endsWith("/returns"))
    expect(writes).toHaveLength(1)
    const form = writes[0]!.init!.body as FormData
    expect(form.get("baseVersionId")).toBe("v1"); expect(form.get("priorSha256")).toBe(versions[0]!.sha256)
    expect((form.get("file") as File).name).toBe("Renamed return.docx"); expect(refreshed).toBe(1)
    const comparisons = view.requests.filter(request => request.url.pathname.endsWith("/compare"))
    expect(comparisons).toHaveLength(2)
    for (const request of comparisons) { expect(request.url.searchParams.get("leftVersion")).toBe("v1"); expect(request.url.searchParams.get("rightVersion")).toBe("returned-v4") }
  } finally { await view.cleanup() }
})

test("an arbitrary return with no suggestion can choose any retained document and older base without trusting its filename", async () => {
  const other = [version("other3", 3), version("other1", 1)]
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }, { name: "Another agreement.docx" }, { name: "Notes.pdf" }], onNew: () => {}, onSaved: () => {} })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection({ kind: "new-document", selected: undefined, candidates: [], reason: "No supported match found." }))
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved("Another agreement.docx", other, "other3"))
    if (request.url.pathname.endsWith("/returns")) return Response.json({ error: "Current copy changed. Refresh before saving." }, { status: 409 })
    throw new Error(`Unexpected UI request: ${request.method} ${request.url.pathname}`)
  })
  try {
    await view.upload(); expect(Boolean(view.button("Save returned version and compare"))).toBe(false)
    expect(view.field<HTMLSelectElement>("Document")!.value).toBe("")
    expect([...view.field<HTMLSelectElement>("Document")!.options].map(option => option.value)).not.toContain("Notes.pdf")
    await view.select("Document", "Another agreement.docx"); await view.select("Base version", "other1")
    await view.click("Save returned version and compare")
    const form = view.requests.find(request => request.url.pathname.endsWith("/returns"))!.init!.body as FormData
    expect(form.get("baseVersionId")).toBe("other1"); expect(form.get("priorSha256")).toBe(other[0]!.sha256)
    expect(view.text()).toContain("Current copy changed. Refresh before saving.")
    expect(Boolean(view.container.querySelector("dialog"))).toBe(true)
  } finally { await view.cleanup() }
})

test("an exact duplicate is not imported accidentally and can deliberately be kept as a separate document", async () => {
  const added: File[] = []
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }], onNew: file => added.push(file), onSaved: () => { throw new Error("duplicate must not save") } })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection({ kind: "duplicate", sha256: candidate.sha256 }))
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved())
    throw new Error(`Unexpected duplicate write: ${request.method} ${request.url.pathname}`)
  })
  try {
    const file = await view.upload()
    expect(view.text()).toContain("This copy is already kept")
    expect(Boolean(view.button("Save returned version and compare"))).toBe(false)
    await view.select("Base version", "v2")
    expect(view.text()).toContain("This copy is already kept")
    expect(Boolean(view.button("Save returned version and compare"))).toBe(false)
    await view.click("Add as separate document")
    expect(added).toEqual([file]); expect(Boolean(view.container.querySelector("dialog"))).toBe(false)
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns"))).toHaveLength(0)
  } finally { await view.cleanup() }
})

test("switching documents during a pending send never closes or reuses the next document's dialog", async () => {
  let finish!: (response: Response) => void
  const { default: DocumentExchanges } = await import("../src/components/DocumentExchanges")
  const view = await render(async () => createElement(DocumentExchanges, { matterId: "matter", name: "First.docx", versions }), request => {
    if (request.url.pathname.endsWith("/preferences")) return Response.json({ documentMatchingId: false })
    if (request.url.pathname.endsWith("/exchanges")) return Response.json({ exchanges: [] })
    if (request.url.pathname.endsWith("/send")) return new Promise(resolve => { finish = resolve })
    throw new Error(`Unexpected request: ${request.url.pathname}`)
  })
  try {
    await view.click("Record a send"); await view.click("Record as sent and download")
    const other = [version("other3", 3), version("other1", 1)]
    await view.update(createElement(DocumentExchanges, { matterId: "matter", name: "Second.docx", versions: other }))
    // Keeping the fresh dialog open is permitted; retaining the old selection is not.
    const remainedOpen = Boolean(view.container.querySelector("dialog"))
    if (remainedOpen) expect(view.field<HTMLSelectElement>("Version")?.value).toBe("other3")
    await act(async () => { finish(new Response("exact first copy")) })
    if (remainedOpen) {
      expect(view.container.querySelector("dialog")?.textContent).toContain("Second.docx")
      expect(view.field<HTMLSelectElement>("Version")?.value).toBe("other3")
    }
    const writes = view.requests.filter(request => request.method === "POST")
    expect(writes).toHaveLength(1); expect(writes[0]!.url.pathname).toContain("/versions/v3/send")
  } finally { await view.cleanup() }
})

test("a late automatic candidate version response cannot overwrite an explicit different document choice", async () => {
  let finish!: (response: Response) => void
  const other = [version("other3", 3), version("other1", 1)]
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }, { name: "Another agreement.docx" }], onNew: () => {}, onSaved: () => {} })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection())
    if (request.url.pathname.endsWith("/versions")) return request.url.searchParams.get("name") === "Agreement.docx" ? new Promise(resolve => { finish = resolve }) : Response.json(saved("Another agreement.docx", other, "other3"))
    if (request.url.pathname.endsWith("/returns")) return Response.json({ error: "Synthetic refusal after input verification." }, { status: 409 })
    throw new Error(`Unexpected request: ${request.url.pathname}`)
  })
  try {
    await view.upload(); await view.select("Document", "Another agreement.docx")
    await act(async () => { finish(Response.json(saved())) })
    expect(view.field<HTMLSelectElement>("Document")!.value).toBe("Another agreement.docx")
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("other3")
    await view.select("Base version", "other1"); await view.click("Save returned version and compare")
    const form = view.requests.find(request => request.url.pathname.endsWith("/returns"))!.init!.body as FormData
    expect(form.get("baseVersionId")).toBe("other1"); expect(form.get("priorSha256")).toBe(other[0]!.sha256)
  } finally { await view.cleanup() }
})

test("an externally changed current copy blocks importing a return without discarding it", async () => {
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }], onNew: () => {}, onSaved: () => {} })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection())
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved("Agreement.docx", versions, null))
    throw new Error(`Unexpected blocked return write: ${request.method} ${request.url.pathname}`)
  })
  try {
    await view.upload()
    expect(view.text()).toContain("current file changed outside its saved version")
    expect(view.button("Save returned version and compare")?.disabled).toBe(true)
    expect(view.button("Add as separate document")?.disabled).toBe(false)
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns"))).toHaveLength(0)
  } finally { await view.cleanup() }
})

test("a completed return cannot reopen its comparison after changing matter", async () => {
  let finish!: (response: Response) => void, refreshed = 0
  const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
  const props = { documents: [{ name: "Agreement.docx" }], onNew: () => {}, onSaved: () => { refreshed++ } }
  const view = await render(async () => createElement(ReturnedDocument, { ...props, matterId: "first" }), request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection())
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved())
    if (request.url.pathname.endsWith("/returns")) return new Promise(resolve => { finish = resolve })
    throw new Error(`An abandoned return opened an unexpected view: ${request.url.pathname}`)
  })
  try {
    await view.upload(); await view.click("Save returned version and compare")
    await view.update(createElement(ReturnedDocument, { ...props, matterId: "second" }))
    expect(Boolean(view.container.querySelector("dialog"))).toBe(false)
    await act(async () => { finish(Response.json({ name: "Agreement.docx", version: { id: "old-return", number: 4 }, comparison: { leftVersion: "v1", rightVersion: "old-return" } })) })
    expect(refreshed).toBe(0); expect(Boolean(view.container.querySelector('[role="dialog"]'))).toBe(false)
    expect(view.requests.some(request => request.url.pathname.endsWith("/compare"))).toBe(false)
  } finally { await view.cleanup() }
})

test("authenticated email inspection preserves explicit base and recipient choices and removes stale source on rejection", async () => {
  let finishEmail!: (response: Response) => void
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }, { name: "Client.eml" }, { name: "Other.msg" }, { name: "Wrong.eml" }], onNew: () => {}, onSaved: () => {} })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) {
      const email = (request.init!.body as FormData).get("emailName")
      if (email === "Client.eml") return new Promise(resolve => { finishEmail = resolve })
      if (email === "Other.msg") return Response.json(inspection({ selected: { ...candidate, versionId: "v3", versionNumber: 3 }, source: { documentName: "Other.msg", attachmentId: "part:2", sender: "Other sender" } }))
      if (email === "Wrong.eml") return Response.json({ error: "This file is not an exact attachment of the chosen email." }, { status: 409 })
      return Response.json(inspection())
    }
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved())
    if (request.url.pathname.endsWith("/returns")) return Response.json({ error: "Synthetic stop after validating submitted source." }, { status: 409 })
    throw new Error(`Unexpected email return request: ${request.url.pathname}`)
  })
  try {
    await view.upload(); await view.select("Base version", "v2")
    await view.select("Email source", "Client.eml")
    expect(Boolean(view.button("Save returned version and compare"))).toBe(false)
    expect(Boolean(view.container.querySelector('[aria-label="Recorded email source"]'))).toBe(false)
    await act(async () => { finishEmail(Response.json(inspection({ selected: { ...candidate, versionId: "v3", versionNumber: 3 }, source: { documentName: "Client.eml", attachmentId: "part:1", sender: "Client Sender", date: "2026-11-09T09:00:00Z" } }))) })
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("v2")
    expect(view.field<HTMLInputElement>("Received from")!.value).toBe("Client Sender")
    expect(view.text()).toContain("Date: 2026-11-09T09:00:00Z")
    await view.input("Received from", "Human selected party"); await view.select("Email source", "Other.msg")
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("v2")
    expect(view.field<HTMLInputElement>("Received from")!.value).toBe("Human selected party")
    expect(view.text()).toContain("Sender: Other sender"); expect(view.text()).toContain("Date: Not recorded")
    await view.click("Save returned version and compare")
    const form = view.requests.find(request => request.url.pathname.endsWith("/returns"))!.init!.body as FormData
    expect(form.get("emailName")).toBe("Other.msg"); expect(form.get("baseVersionId")).toBe("v2"); expect(form.get("party")).toBe("Human selected party")
    expect(form.has("sender")).toBe(false); expect(form.has("date")).toBe(false); expect(form.has("attachmentId")).toBe(false)
    await view.select("Email source", "Wrong.eml")
    expect(view.text()).toContain("not an exact attachment")
    expect(Boolean(view.container.querySelector('[aria-label="Recorded email source"]'))).toBe(false)
    expect(Boolean(view.button("Save returned version and compare"))).toBe(false)
  } finally { await view.cleanup() }
})

test("a normal renamed upload offers return matching once; explicit keep separate bypasses reinspection", async () => {
  const xhrBefore = Object.getOwnPropertyDescriptor(globalThis, "XMLHttpRequest")
  const forms: FormData[] = []; let refreshed = 0
  const incoming = [new File(["reviewed synthetic return"], "Renamed.docx")]
  const documents: Document[] = [{ id: 1, name: "Agreement.docx", doc_path: "/matter/Agreement.docx", created_at: 1, role: "source" }]
  Object.defineProperty(globalThis, "XMLHttpRequest", { configurable: true, writable: true, value: class {
    upload: { onload?: () => void } = {}; onload?: () => void; timeout = 0
    status = 200; responseText = JSON.stringify({ name: "Renamed.docx", docPath: "/matter/Renamed.docx", sections: 1, chunks: 1, injection: null })
    open(method: string, url: string) { expect(method).toBe("POST"); expect(url).toContain("/matters/matter/documents") }
    send(form: FormData) { forms.push(form); this.upload.onload?.(); this.onload?.() }
  } })
  const view = await render(async () => {
    const { default: DocumentUpload } = await import("../src/components/DocumentUpload")
    return createElement(DocumentUpload, { matterId: "matter", documents, incoming, onIncoming: () => {}, onView: () => {}, onUploaded: () => { refreshed++ } })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(inspection())
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved())
    throw new Error(`Unexpected normal upload request: ${request.url.pathname}`)
  })
  try {
    expect(Boolean(view.container.querySelector('dialog[aria-label="Review returned file"]'))).toBe(true)
    expect(forms).toHaveLength(0)
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns/inspect"))).toHaveLength(1)
    await view.click("Add as separate document")
    expect(forms).toHaveLength(1); expect(forms[0]!.get("existing")).toBe("keep-both")
    expect((forms[0]!.get("file") as File).name).toBe("Renamed.docx"); expect(refreshed).toBe(1)
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns/inspect"))).toHaveLength(1)
    expect(Boolean(view.container.querySelector("dialog"))).toBe(false)
  } finally {
    await view.cleanup()
    if (xhrBefore) Object.defineProperty(globalThis, "XMLHttpRequest", xhrBefore); else delete (globalThis as Record<string, unknown>).XMLHttpRequest
  }
})

test("return identification advice is explicit, never changes the base automatically, and is not sent as save authority", async () => {
  const other = [version("other3", 3), version("other1", 1)]
  const alternative: ReturnCandidate = { ...candidate, documentId: "other-doc", documentName: "Another agreement.docx", versionId: "other1", versionNumber: 1, sha256: other[1]!.sha256 }
  const uncertain = inspection({ kind: "needs-confirmation", selected: undefined, candidates: [candidate, alternative], reason: "Several candidates need confirmation." })
  const advisory = { choice: "candidate", versionId: "other1", reason: "Distinctive matching clause wording.", evidence: [{ versionId: "other1", incomingQuote: "A distinctive term", candidateQuote: "A distinctive term" }], model: { providerID: "fixture", id: "reviewer", variant: "bounded" }, incomingHash: uncertain.sha256, candidates: uncertain.candidates.map(item => ({ versionId: item.versionId, sha256: item.sha256 })), limitations: ["A suggestion does not establish document ancestry."] }
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }, { name: "Another agreement.docx" }], onNew: () => {}, onSaved: () => {} })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(uncertain)
    if (request.url.pathname.endsWith("/returns/tie-break")) return Response.json({ inspection: uncertain, advisory })
    if (request.url.pathname.endsWith("/versions")) return Response.json(request.url.searchParams.get("name") === "Agreement.docx" ? saved() : saved("Another agreement.docx", other, "other3"))
    if (request.url.pathname.endsWith("/returns")) return Response.json({ error: "Synthetic stop after confirming the chosen base." }, { status: 409 })
    throw new Error(`Unexpected advice request: ${request.url.pathname}`)
  })
  try {
    await view.upload(); await view.select("Base version", "v2")
    expect(view.requests.filter(request => request.url.pathname.endsWith("/tie-break"))).toHaveLength(0)
    await view.click("Help identify this return")
    expect(view.text()).toContain("Model: fixture/reviewer · bounded")
    expect(view.text()).toContain("Returned file: A distinctive term")
    expect(view.field<HTMLSelectElement>("Document")!.value).toBe("Agreement.docx")
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("v2")
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns"))).toHaveLength(0)
    await view.click("Use suggested base")
    expect(view.field<HTMLSelectElement>("Document")!.value).toBe("Another agreement.docx")
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("other1")
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns"))).toHaveLength(0)
    await view.click("Save returned version and compare")
    const form = view.requests.find(request => request.url.pathname.endsWith("/returns"))!.init!.body as FormData
    expect([...form.keys()].sort()).toEqual(["baseVersionId", "file", "priorSha256", "requestId"])
    expect(form.get("baseVersionId")).toBe("other1"); expect(form.get("priorSha256")).toBe(other[0]!.sha256)
  } finally { await view.cleanup() }
})

test("changing email context aborts a pending tie-break and its late advice cannot repaint", async () => {
  const uncertain = inspection({ kind: "needs-confirmation", selected: undefined, candidates: [candidate, { ...candidate, versionId: "v2", versionNumber: 2 }] })
  let finish!: (response: Response) => void, signal: AbortSignal | null | undefined
  const view = await render(async () => {
    const { default: ReturnedDocument } = await import("../src/components/ReturnedDocument")
    return createElement(ReturnedDocument, { matterId: "matter", documents: [{ name: "Agreement.docx" }, { name: "Client.eml" }], onNew: () => {}, onSaved: () => {} })
  }, request => {
    if (request.url.pathname.endsWith("/returns/inspect")) return Response.json(uncertain)
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved())
    if (request.url.pathname.endsWith("/returns/tie-break")) { signal = request.init?.signal; return new Promise(resolve => { finish = resolve }) }
    throw new Error(`Unexpected stale advice request: ${request.url.pathname}`)
  })
  try {
    await view.upload(); await view.click("Help identify this return")
    await view.select("Email source", "Client.eml")
    expect(signal?.aborted).toBe(true)
    await act(async () => { finish(Response.json({ inspection: uncertain, advisory: { choice: "uncertain", reason: "Old email advice must stay discarded.", evidence: [], model: { providerID: "fixture", id: "reviewer" }, incomingHash: uncertain.sha256, candidates: [], limitations: [] } })) })
    expect(view.text()).not.toContain("Old email advice must stay discarded")
    expect(Boolean(view.container.querySelector('[aria-label="Return identification advice"]'))).toBe(false)
    expect(view.requests.filter(request => request.url.pathname.endsWith("/returns"))).toHaveLength(0)
  } finally { await view.cleanup() }
})

test("correcting a linked returned document sends the exact old association and retains manual choice on a conflict", async () => {
  const exchange: DocumentExchange = { id: 44, kind: "received", documentId: "physical-source", versionId: "physical-v77", party: "Reviewer", actor: "Avery", occurredAt: 1, recordedAt: 1, baseVersionId: "original-base", requestId: "received-once", payload: null, association: { documentId: "old-target", baseVersionId: "linked-base" }, linkedBranch: { exchangeId: 44, sourceDocumentId: "physical-source", sourceName: "Separate return.docx", sourceCurrentVersionId: "physical-v77", versionId: "physical-v77", versionNumber: 77, sha256: "f".repeat(64), ext: ".docx", bytes: 100, targetDocumentId: "old-target", targetName: "Agreement.docx", targetCurrentVersionId: "v3", baseVersionId: "linked-base", baseVersionNumber: 1, current: false, available: true, unavailableReason: null } }
  const view = await render(async () => {
    const { default: DocumentExchanges } = await import("../src/components/DocumentExchanges")
    return createElement(DocumentExchanges, { matterId: "matter", name: "Agreement.docx", versions })
  }, request => {
    if (request.url.pathname.endsWith("/exchanges")) return Response.json({ exchanges: [exchange] })
    if (request.url.pathname === "/matters/matter") return Response.json({ documents: [{ name: "New target.docx" }, { name: "Notes.pdf" }] })
    if (request.url.pathname.endsWith("/versions")) return Response.json(saved("New target.docx", [version("t3", 3), version("t1", 1)], "t3"))
    if (request.url.pathname.endsWith("/44/document")) return Response.json({ error: "This receipt relationship changed. Refresh it before choosing again." }, { status: 409 })
    throw new Error(`Unexpected link correction request: ${request.url.pathname}`)
  })
  try {
    await view.click("Correct linked document")
    expect(view.button("Confirm document link")?.disabled).toBe(true)
    expect([...view.field<HTMLSelectElement>("Document")!.options].map(option => option.value)).toEqual(["", "New target.docx"])
    await view.select("Document", "New target.docx"); await view.select("Base version", "t1")
    await view.click("Confirm document link")
    expect(view.text()).toContain("This receipt relationship changed")
    expect(view.field<HTMLSelectElement>("Document")!.value).toBe("New target.docx")
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("t1")
    await view.click("Confirm document link")
    const writes = view.requests.filter(request => request.method === "POST")
    expect(writes).toHaveLength(2); expect(writes[0]!.init!.body).toBe(writes[1]!.init!.body)
    expect(JSON.parse(writes[0]!.init!.body as string)).toMatchObject({ targetDocumentId: "doc:New target.docx", baseVersionId: "t1", expectedDocumentId: "old-target", expectedBaseVersionId: "linked-base" })
  } finally { await view.cleanup() }
})

test("base correction excludes the received version and keeps its request identity and choice after a server conflict", async () => {
  const exchange: DocumentExchange = { id: 45, kind: "received", documentId: "doc", versionId: "v3", party: null, actor: "Avery", occurredAt: 1, recordedAt: 1, baseVersionId: "v1", requestId: "receipt-one", payload: null }
  const view = await render(async () => {
    const { default: DocumentExchanges } = await import("../src/components/DocumentExchanges")
    return createElement(DocumentExchanges, { matterId: "matter", name: "Agreement.docx", versions })
  }, request => {
    if (request.url.pathname.endsWith("/exchanges")) return Response.json({ exchanges: [exchange] })
    if (request.url.pathname.endsWith("/45/base")) return Response.json({ error: "The recorded base changed. Refresh this receipt." }, { status: 409 })
    throw new Error(`Unexpected base correction request: ${request.url.pathname}`)
  })
  try {
    await view.click("Correct base version")
    expect([...view.field<HTMLSelectElement>("Base version")!.options].map(option => option.value)).toEqual(["", "v2", "v1"])
    expect(view.button("Confirm base correction")?.disabled).toBe(true)
    await view.select("Base version", "v2"); await view.click("Confirm base correction")
    expect(view.field<HTMLSelectElement>("Base version")!.value).toBe("v2")
    expect(view.text()).toContain("The recorded base changed")
    await view.click("Confirm base correction")
    const writes = view.requests.filter(request => request.method === "POST")
    expect(writes).toHaveLength(2); expect(writes[0]!.init!.body).toBe(writes[1]!.init!.body)
    expect(JSON.parse(writes[0]!.init!.body as string)).toMatchObject({ baseVersionId: "v2", expectedBaseVersionId: "v1" })
  } finally { await view.cleanup() }
})

test("linked return display preserves physical version numbering and never marks it as the target current copy", async () => {
  const branch: LinkedDocumentBranch = { exchangeId: 44, sourceDocumentId: "physical-source", sourceName: "Separate return.docx", sourceCurrentVersionId: "physical-v77", versionId: "physical-v77", versionNumber: 77, sha256: "f".repeat(64), ext: ".docx", bytes: 100, targetDocumentId: "doc:Agreement.docx", targetName: "Agreement.docx", targetCurrentVersionId: "v3", baseVersionId: "v1", baseVersionNumber: 1, current: false, available: true, unavailableReason: null }
  const compared: unknown[] = [], props = { matterId: "matter", name: "Agreement.docx", unavailable: {}, onCompare: (input: unknown) => compared.push(input), onRestored: () => {}, onUnavailable: () => {} }
  const { default: VersionsPanel } = await import("../src/components/DocumentVersions")
  const view = await render(async () => createElement(VersionsPanel, { ...props, data: { ...saved(), linked_returns: [branch] } }), request => {
    if (request.url.pathname.endsWith("/exchanges")) return Response.json({ exchanges: [] })
    throw new Error(`Unexpected linked-branch request: ${request.url.pathname}`)
  })
  try {
    expect(view.container.querySelectorAll(".versions-list .version-row")).toHaveLength(3)
    expect(view.container.querySelector(".versions-list")?.textContent).not.toContain("Version 77")
    const linked = view.container.querySelector('[aria-label="Linked returned branches"]')!
    expect(linked.textContent).toContain("Separate return.docx · Version 77")
    expect(linked.textContent).toContain("Based on Agreement.docx · Version 1")
    expect(Boolean(linked.querySelector(".version-tag"))).toBe(false)
    expect(linked.querySelector("a")?.getAttribute("href")).toContain("/exchanges/44/branch/content")
    await view.click("Compare with linked base")
    expect(compared).toEqual([{ left: "Agreement.docx", right: "Separate return.docx", leftVersion: "v1", rightVersion: "physical-v77" }])
    await view.update(createElement(VersionsPanel, { ...props, data: { ...saved(), linked_returns: [{ ...branch, available: false, unavailableReason: "source-removed" }] } }))
    expect(Boolean(view.container.querySelector('[aria-label="Linked returned branches"] a'))).toBe(false)
    expect(Boolean(view.button("Compare with linked base"))).toBe(false)
    expect(view.text()).toContain("linked return or its base is no longer available")
  } finally { await view.cleanup() }
})
