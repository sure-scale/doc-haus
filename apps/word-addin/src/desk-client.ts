import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { WORD_UPDATE_MESSAGE, httpsOrigin } from "./client"
import type { DeskRequest } from "./desk-protocol"
import { captureSelection, showSelection } from "./selection"
import { DeskRequestError } from "./desk-host"

/** The pane keeps the capability. The dialog receives only scoped results. */
export function deskRequestHandler(
  link: WordDocumentLink,
  request: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = (input, init) => globalThis.fetch(input, init),
  timeoutMs = 60_000,
): (input: DeskRequest) => Promise<unknown> {
  const origin = httpsOrigin(link.apiOrigin)
  return async (input) => {
    if (input.action === "selection") return captureSelection()
    if (input.action === "show") {
      await showSelection(input.selection)
      return null
    }
    const abort = new AbortController()
    const timeout = setTimeout(() => abort.abort(), timeoutMs)
    try {
      const { action, ...body } = input
      const response = await request(`${origin}/word/desk/${action}`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${link.token}`,
          "x-dochaus-word-version": WORD_ADDIN_VERSION,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
        credentials: "omit",
        redirect: "error",
        referrerPolicy: "no-referrer",
        signal: abort.signal,
      })
      if (!response.ok) {
        if (response.status === 426) throw new DeskRequestError(WORD_UPDATE_MESSAGE)
        const value = await response.json().catch(() => null) as { message?: unknown; error?: unknown } | null
        const detail = typeof value?.message === "string" ? value.message : typeof value?.error === "string" ? value.error : undefined
        throw new DeskRequestError(detail ?? "Matter documents are unavailable. Check the document link and try again.")
      }
      return await response.json()
    } catch (error) {
      if (abort.signal.aborted) throw new DeskRequestError("doc.haus did not respond in time. Try the action again.")
      if (error instanceof TypeError) throw new DeskRequestError("Cannot reach doc.haus. Check that it is running and try again.")
      throw error
    } finally {
      clearTimeout(timeout)
    }
  }
}
