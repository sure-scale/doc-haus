import { WORD_PROTOCOL_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { WORD_UPDATE_MESSAGE, httpsOrigin, reloadAfterUpdate } from "./client"
import { isDeskRequest, type DeskRequest } from "./desk-protocol"
import { captureSelection, showSelection } from "./selection"
import { DeskRequestError } from "./desk-host"

/** The pane keeps the capability. The dialog receives only scoped results. */
export function deskRequestHandler(
  link: WordDocumentLink,
  request: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = (input, init) => globalThis.fetch(input, init),
  timeoutMs = 60_000,
): (input: DeskRequest) => Promise<unknown> {
  const origin = httpsOrigin(link.apiOrigin)
  const browser = new URL(link.origin)
  if (!/^https?:$/.test(browser.protocol) || browser.username || browser.password || browser.href !== `${browser.origin}/`)
    throw new Error("A doc.haus browser origin is required.")
  const matterUrl = new URL(`/matter/${encodeURIComponent(link.matter)}?view=documents`, browser.origin).href
  return async (input) => {
    if (!isDeskRequest(input)) throw new DeskRequestError("Choose a supported document up to 16 MB.")
    if (input.action === "chat-materials" || input.action === "chat-attach" || input.action === "chat-upload")
      throw new DeskRequestError("Open Attach from the chat to choose material.")
    if (input.action === "selection") return captureSelection()
    if (input.action === "show") {
      await showSelection(input.selection)
      return null
    }
    const abort = new AbortController()
    const timeout = setTimeout(() => abort.abort(), timeoutMs)
    try {
      const { action, ...body } = input
      let upload: FormData | undefined
      if (input.action === "document-upload") {
        const binary = atob(input.data)
        const bytes = Uint8Array.from(binary, character => character.charCodeAt(0))
        upload = new FormData()
        upload.set("file", new File([bytes], input.name))
      }
      const response = await request(`${origin}/word/desk/${upload ? "upload" : action}`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${link.token}`,
          "x-dochaus-word-version": WORD_PROTOCOL_VERSION,
          ...(!upload ? { "Content-Type": "application/json" } : {}),
        },
        body: upload ?? JSON.stringify(input.action === "open" ? { ...body, origin: browser.origin } : body),
        credentials: "omit",
        redirect: "error",
        referrerPolicy: "no-referrer",
        signal: abort.signal,
      })
      if (!response.ok) {
        if (response.status === 426) { await reloadAfterUpdate(); throw new DeskRequestError(WORD_UPDATE_MESSAGE) }
        const value = await response.json().catch(() => null) as { message?: unknown; error?: unknown } | null
        const detail = typeof value?.message === "string" ? value.message : typeof value?.error === "string" ? value.error : undefined
        throw new DeskRequestError(detail ?? "Matter documents are unavailable. Check the document link and try again.")
      }
      const value = await response.json()
      if (input.action === "sources") return { ...value, matterUrl }
      if (input.action === "open") {
        try {
          const uri = new URL(typeof value?.uri === "string" ? value.uri : "")
          const source = new URL(uri.searchParams.get("source") ?? "")
          if (uri.protocol !== "doc.haus:" || uri.host !== "word" || uri.pathname !== "/open" || uri.username || uri.password || uri.hash || [...uri.searchParams].length !== 1 || source.origin !== origin || source.username || source.password || source.search || source.hash || !/^\/word\/open\/[A-Za-z0-9_-]{43}$/.test(source.pathname) || typeof value.id !== "string" || !/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(value.id)) throw new Error()
          return { uri: uri.href, id: value.id }
        } catch { throw new DeskRequestError("The Word connection could not be prepared. Keep doc.haus running and try again.") }
      }
      return value
    } catch (error) {
      if (abort.signal.aborted) throw new DeskRequestError("doc.haus did not respond in time. Try the action again.")
      if (error instanceof TypeError) throw new DeskRequestError("Cannot reach doc.haus. Check that it is running and try again.")
      throw error
    } finally {
      clearTimeout(timeout)
    }
  }
}
