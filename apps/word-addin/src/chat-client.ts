import { WORD_PROTOCOL_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { WORD_CHAT_ROUTES } from "../../../dochaus/lib/word-chat-contract"
import { httpsOrigin, reloadAfterUpdate, WORD_UPDATE_MESSAGE } from "./client"
import { OpenCode, type V2Event } from "@opencode/client"

export class ChatRequestError extends Error {
  constructor(message: string, readonly retryable = false) { super(message); this.name = "ChatRequestError" }
}
export type ChatAction = (typeof WORD_CHAT_ROUTES)[number]
export interface WordChatApi {
  request<T>(action: ChatAction, input?: object, signal?: AbortSignal): Promise<T>
  upload<T>(conversationId: string, file: File, mode: "request" | "matter"): Promise<T>
  events?(conversationId: string, signal: AbortSignal): AsyncIterable<V2Event>
}

export function chatClient(
  link: WordDocumentLink,
  request: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = (input, init) => globalThis.fetch(input, init),
  timeoutMs = 60_000,
  prepareHeaders?: () => Record<string, string>,
): WordChatApi {
  const origin = httpsOrigin(link.apiOrigin)
  async function post<T>(action: ChatAction, body: object | FormData, signal?: AbortSignal): Promise<T> {
    const abort = new AbortController()
    const timeout = setTimeout(() => abort.abort(), timeoutMs)
    const cancelled = () => abort.abort()
    if (signal?.aborted) abort.abort()
    signal?.addEventListener("abort", cancelled, { once: true })
    try {
      const multipart = body instanceof FormData
      const response = await request(`${origin}/word/chat/${action}`, {
        method: "POST", headers: {
          Authorization: `Bearer ${link.token}`,
          "x-dochaus-word-version": WORD_PROTOCOL_VERSION,
          ...(action === "prepare" ? prepareHeaders?.() : {}),
          ...(multipart ? {} : { "Content-Type": "application/json" }),
        },
        body: multipart ? body : JSON.stringify(body), credentials: "omit", redirect: "error",
        referrerPolicy: "no-referrer", signal: abort.signal,
      })
      if (!response.ok) {
        if (response.status === 426) { await reloadAfterUpdate(); throw new ChatRequestError(WORD_UPDATE_MESSAGE) }
        const value = await response.json().catch(() => undefined) as { error?: unknown } | undefined
        throw new ChatRequestError(typeof value?.error === "string" ? value.error : "Chat could not complete the request. Keep your message and reconnect.", response.status >= 500 || response.status === 429)
      }
      return await response.json() as T
    } catch (error) {
      if (error instanceof ChatRequestError) throw error
      if (signal?.aborted) throw new ChatRequestError("The connection was interrupted. Reconnect to check the request.", true)
      if (abort.signal.aborted) throw new ChatRequestError("doc.haus did not respond in time. Reconnect to check the request.", true)
      throw new ChatRequestError("Cannot reach doc.haus. Keep it running, then reconnect.", true)
    } finally { clearTimeout(timeout); signal?.removeEventListener("abort", cancelled) }
  }
  return {
    request: <T>(action: ChatAction, input: object = {}, signal?: AbortSignal) => post<T>(action, input, signal),
    upload: <T>(conversationId: string, file: File, mode: "request" | "matter") => {
      const body = new FormData(); body.set("conversationId", conversationId); body.set("file", file); body.set("mode", mode)
      return post<T>("upload", body)
    },
    events(conversationId, signal) {
      // Reuse the native client's streaming parser and socket lifecycle. Its
      // sole endpoint is mapped to the document-granted, content-free proxy.
      const transport = async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = new URL(input instanceof Request ? input.url : String(input))
        if (url.origin !== origin || url.pathname !== "/api/event") throw new ChatRequestError("This chat stream is unavailable.")
        const response = await request(`${origin}/word/chat/events`, {
          method: "POST", headers: { Authorization: `Bearer ${link.token}`, "x-dochaus-word-version": WORD_PROTOCOL_VERSION, "Content-Type": "application/json" },
          body: JSON.stringify({ conversationId }), credentials: "omit", redirect: "error", referrerPolicy: "no-referrer", signal: init?.signal,
        })
        if (response.status === 426) { await response.body?.cancel(); await reloadAfterUpdate(); throw new ChatRequestError(WORD_UPDATE_MESSAGE) }
        return response
      }
      const native = OpenCode.make({ baseUrl: origin, fetch: transport as typeof globalThis.fetch })
      return {
        async *[Symbol.asyncIterator]() {
          for await (const event of native.event.subscribe({ signal })) {
            // Do not render or even forward raw model JSON, reasoning, tools,
            // reviewer text, or unchecked proposal/citation data.
            if (event.type === "server.connected") yield { id: event.id, type: event.type, data: {} }
            else if (event.type === "session.status" && typeof event.data?.sessionID === "string" && ["busy", "idle", "retry"].includes(event.data.status?.type))
              yield { id: event.id, created: event.created, type: event.type, data: { sessionID: event.data.sessionID, status: { type: event.data.status.type === "idle" ? "idle" : "busy" } } }
            else if (event.type === "session.deleted" && typeof event.data?.sessionID === "string")
              yield { id: event.id, created: event.created, type: event.type, data: { sessionID: event.data.sessionID }, durable: event.durable }
          }
        },
      }
    },
  }
}
