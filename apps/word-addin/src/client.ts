import type { WordDocumentLink, WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"

export class ConnectionError extends Error {
  constructor(
    message: string,
    readonly retryable = false,
  ) {
    super(message)
  }
}

export function httpsOrigin(value: string): string {
  const url = new URL(value)
  if (url.protocol !== "https:" || url.username || url.password || url.href !== `${url.origin}/`) {
    throw new Error("A trusted HTTPS address is required.")
  }
  return url.origin
}

export function readLink(value: unknown, paneOrigin: string): WordDocumentLink | null {
  if (typeof value === "string") {
    try {
      value = JSON.parse(value)
    } catch {
      return null
    }
  }
  if (!value || typeof value !== "object") return null
  const link = value as Partial<WordDocumentLink>
  if (
    link.version !== 1 ||
    !["origin", "apiOrigin", "matter", "matterName", "document", "documentName", "baseVersion", "token"].every(
      (key) => typeof link[key as keyof WordDocumentLink] === "string" && Boolean(link[key as keyof WordDocumentLink]),
    )
  ) {
    return null
  }
  try {
    // The instance owns both this pane and the receiver. Never send a document's
    // capability to a different host selected by an altered settings property.
    if (httpsOrigin(link.apiOrigin!) !== httpsOrigin(paneOrigin)) return null
    const browser = new URL(link.origin!)
    if (
      !["https:", "http:"].includes(browser.protocol) ||
      browser.username ||
      browser.password ||
      browser.href !== `${browser.origin}/`
    )
      return null
    return link as WordDocumentLink
  } catch {
    return null
  }
}

export class WordClient {
  private readonly receiver: string
  constructor(
    private readonly link: WordDocumentLink,
    private readonly session: string,
    private readonly request: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = fetch,
    private readonly timeoutMs = 30_000,
  ) {
    this.receiver = httpsOrigin(link.apiOrigin)
  }

  private async post(path: "saves" | "heartbeat" | "done", bytes?: Uint8Array, base?: string): Promise<unknown> {
    const abort = new AbortController()
    const timeout = setTimeout(() => abort.abort(), this.timeoutMs)
    try {
      const response = await this.request(`${this.receiver}/word/${path}`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${this.link.token}`,
          "x-dochaus-word-session": this.session,
          ...(bytes
            ? {
                "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "x-dochaus-word-base-version": base!,
              }
            : {}),
        },
        body: bytes ? new Blob([bytes as Uint8Array<ArrayBuffer>]) : undefined,
        credentials: "omit",
        redirect: "error",
        referrerPolicy: "no-referrer",
        signal: abort.signal,
      })
      if (!response.ok) {
        void response.body?.cancel().catch(() => {})
        if ([401, 403, 404, 410].includes(response.status)) {
          throw new ConnectionError("This document's link is no longer available. Open a new copy from doc.haus.")
        }
        if (response.status === 413)
          throw new ConnectionError("This document is too large to save. Keep the Word file and return to doc.haus.")
        throw new ConnectionError(
          "doc.haus could not accept the document. Try Save to doc.haus again.",
          response.status >= 500 || response.status === 429,
        )
      }
      if (path === "saves") {
        try {
          // Keep the deadline alive through the body, not only response headers.
          return await response.json()
        } catch {
          throw new ConnectionError("doc.haus did not confirm the save. Try Save to doc.haus again.", true)
        }
      }
      // Heartbeat/Done acknowledge through the status; release their unused body.
      void response.body?.cancel().catch(() => {})
    } catch (error) {
      if (error instanceof ConnectionError) throw error
      throw new ConnectionError(
        "Cannot reach doc.haus. Check that it is running, then try Save to doc.haus again.",
        true,
      )
    } finally {
      clearTimeout(timeout)
    }
  }

  async save(bytes: Uint8Array, base: string): Promise<WordSaveResponse> {
    const value = (await this.post("saves", bytes, base)) as WordSaveResponse
    if (
      !value?.version ||
      typeof value.version.id !== "string" ||
      !value.version.id ||
      !Number.isSafeInteger(value.version.number) ||
      value.version.number < 1 ||
      !(value.version.parent === null || typeof value.version.parent === "string") ||
      typeof value.version.created !== "boolean" ||
      typeof value.overlap !== "boolean"
    ) {
      throw new ConnectionError("doc.haus did not confirm the save. Try Save to doc.haus again.", true)
    }
    return value
  }

  async heartbeat(): Promise<void> {
    await this.post("heartbeat")
  }
  async done(): Promise<void> {
    await this.post("done")
  }
}
