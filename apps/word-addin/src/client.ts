import {
  WORD_ADDIN_VERSION,
  type WordConnectionState,
  type WordLeaseAction,
  type WordDocumentLink,
  type WordSaveResponse,
} from "../../../dochaus/lib/word-addin-contract"

export const WORD_UPDATE_MESSAGE =
  "This Word add-in is out of date. Update Word setup in doc.haus, restart Word, then open a fresh copy."

export type LatestWordContent = { version: { id: string; number: number }; bytes: Uint8Array }

export function isCurrentPaneVersion(value: string): boolean {
  try {
    const versions = new URL(value).searchParams.getAll("version")
    return versions.length === 1 && versions[0] === WORD_ADDIN_VERSION
  } catch {
    return false
  }
}

export class ConnectionError extends Error {
  constructor(
    message: string,
    readonly retryable = false,
    readonly connection?: WordConnectionState,
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

export function readConnection(value: unknown): WordConnectionState | undefined {
  if (!value || typeof value !== "object") return
  const state = value as WordConnectionState
  const ref = (value: unknown) =>
    value === null ||
    (typeof value === "object" &&
      value !== null &&
      typeof (value as { id?: unknown }).id === "string" &&
      Boolean((value as { id?: unknown }).id) &&
      Number.isSafeInteger((value as { number?: unknown }).number) &&
      Number((value as { number?: unknown }).number) > 0)
  if (
    !state.version ||
    !state.lease ||
    (state.nativeOperation !== undefined && (!state.nativeOperation || typeof state.nativeOperation !== "object" || !Number.isSafeInteger(state.nativeOperation.sequence) || state.nativeOperation.sequence < 0 || !["idle", "pending", "confirmed", "not-dispatched"].includes(state.nativeOperation.state))) ||
    !["current", "stale", "removed", "unrecorded-change"].includes(state.version.state) ||
    !ref(state.version.base) ||
    !ref(state.version.current) ||
    !["available", "owned", "other"].includes(state.lease.state) ||
    !(state.lease.expiresAt === null || Number.isFinite(state.lease.expiresAt)) ||
    (state.lease.state !== "available" && (typeof state.lease.generation !== "string" || !state.lease.generation))
  )
    return
  if (state.version.state === "current" && (!state.version.base || state.version.base.id !== state.version.current?.id))
    return
  return state
}

export class WordClient {
  private readonly receiver: string
  constructor(
    private readonly link: WordDocumentLink,
    private readonly session: string,
    private readonly request: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = (input, init) =>
      globalThis.fetch(input, init),
    private readonly timeoutMs = 30_000,
  ) {
    this.receiver = httpsOrigin(link.apiOrigin)
  }

  private async post(
    path: "saves" | "heartbeat" | "done" | "lease" | "latest" | "native-operation",
    options: {
      bytes?: Uint8Array
      base: string
      generation?: string
      input?: object
    },
  ): Promise<unknown> {
    const abort = new AbortController()
    const timeout = setTimeout(() => abort.abort(), this.timeoutMs)
    try {
      const response = await this.request(`${this.receiver}/word/${path}`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${this.link.token}`,
          "x-dochaus-word-session": this.session,
          "x-dochaus-word-version": WORD_ADDIN_VERSION,
          "x-dochaus-word-base-version": options.base,
          ...(options.generation ? { "x-dochaus-word-lease": options.generation } : {}),
          ...(options.bytes
            ? { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
            : options.input
              ? { "Content-Type": "application/json" }
              : {}),
        },
        body: options.bytes
          ? new Blob([options.bytes as Uint8Array<ArrayBuffer>])
          : options.input
            ? JSON.stringify(options.input)
            : undefined,
        credentials: "omit",
        redirect: "error",
        referrerPolicy: "no-referrer",
        signal: abort.signal,
      })
      if (!response.ok) {
        if ([409, 410, 423].includes(response.status)) {
          const value = await response.json().catch(() => undefined)
          const connection = readConnection(value)
          if (connection)
            throw new ConnectionError(
              "Saving is paused. Keep your Word file and check the version and editing status in this pane.",
              false,
              connection,
            )
        }
        void response.body?.cancel().catch(() => {})
        if (response.status === 426) throw new ConnectionError(WORD_UPDATE_MESSAGE)
        if ([401, 403, 404, 410].includes(response.status))
          throw new ConnectionError(
            "This document's link is no longer available. Keep your Word file and open a new copy from doc.haus.",
          )
        if (response.status === 413)
          throw new ConnectionError("This document is too large to save. Keep the Word file and return to doc.haus.")
        throw new ConnectionError(
          "doc.haus could not accept the request. Check editing status, then try again.",
          response.status >= 500 || response.status === 429,
        )
      }
      try {
        // Every command returns authoritative connection state. Keep its entire
        // JSON body inside the deadline, including headers-only stalled replies.
        return await response.json()
      } catch {
        throw new ConnectionError("doc.haus did not confirm the request. Check the connection and try again.", true)
      }
    } catch (error) {
      if (error instanceof ConnectionError) throw error
      throw new ConnectionError("Cannot reach doc.haus. Check that it is running, then try again.", true)
    } finally {
      clearTimeout(timeout)
    }
  }

  private connection(value: unknown): WordConnectionState {
    const state = readConnection(value)
    if (!state)
      throw new ConnectionError("doc.haus did not confirm editing status. Check the connection and try again.", true)
    return state
  }

  async lease(
    action: WordLeaseAction,
    base: string,
    generation?: string,
    expected?: { current: string; generation: string },
  ): Promise<WordConnectionState> {
    return this.connection(
      await this.post("lease", {
        base,
        generation,
        input: {
          action,
          baseVersion: base,
          ...(expected ? { expectedCurrent: expected.current, expectedGeneration: expected.generation } : {}),
        },
      }),
    )
  }

  /** No retry grants permission to dispatch: the caller must know this call
   * has not previously queued a Word mutation. */
  async nativeOperation(base: string, generation: string, input: { action: "read" } | { action: "begin"; sequence: number; operationId: string; intent: "amendment"; proposal: { conversationId: string; turnId: string } } | { action: "begin"; sequence: number; operationId: string; intent: "refresh"; expectedCurrent: string } | { action: "settle"; sequence: number; operationId: string; outcome: "confirmed" | "not-dispatched" }) {
    const value = await this.post("native-operation", { base, generation, input }) as { operation?: { sequence?: unknown; operationId?: unknown; state?: unknown }; connection?: unknown }
    const operation = value?.operation
    if (!operation || !Number.isSafeInteger(operation.sequence) || Number(operation.sequence) < 0 ||
      !(operation.operationId === null || typeof operation.operationId === "string" && /^[A-Za-z0-9_-]{1,128}$/.test(operation.operationId)) ||
      !["idle", "pending", "confirmed", "not-dispatched"].includes(String(operation.state)))
      throw new ConnectionError("doc.haus could not confirm the Word operation. Inspect the document before continuing.", true)
    const connection = this.connection(value.connection)
    if (!connection.nativeOperation || connection.nativeOperation.sequence !== operation.sequence || connection.nativeOperation.state !== operation.state ||
      input.action !== "read" && (operation.operationId !== input.operationId || operation.sequence !== (input.action === "begin" ? input.sequence + 1 : input.sequence)) ||
      input.action === "begin" && (connection.lease.state !== "owned" || connection.lease.generation !== generation || connection.version.base?.id !== base ||
        (input.intent === "refresh" ? !["current", "stale"].includes(connection.version.state) || connection.version.current?.id !== input.expectedCurrent
          : connection.version.state !== "current" || connection.version.current?.id !== base)) ||
      input.action === "settle" && operation.state !== input.outcome)
      throw new ConnectionError("doc.haus could not confirm the Word operation. Inspect the document before continuing.", true)
    return { operation: operation as { sequence: number; operationId: string | null; state: "idle" | "pending" | "confirmed" | "not-dispatched" }, connection }
  }

  async save(bytes: Uint8Array, base: string, generation: string): Promise<WordSaveResponse> {
    const value = (await this.post("saves", { bytes, base, generation })) as WordSaveResponse
    if (
      !value?.version ||
      typeof value.version.id !== "string" ||
      !value.version.id ||
      !Number.isSafeInteger(value.version.number) ||
      value.version.number < 1 ||
      !(value.version.parent === null || typeof value.version.parent === "string") ||
      typeof value.version.created !== "boolean" ||
      typeof value.overlap !== "boolean"
    )
      throw new ConnectionError("doc.haus did not confirm the save. Try Save to doc.haus again.", true)
    const connection = this.connection(value.connection)
    if (
      connection.version.state !== "current" ||
      connection.version.base?.id !== value.version.id ||
      connection.version.current?.id !== value.version.id ||
      connection.lease.state !== "owned" ||
      connection.lease.generation !== generation
    )
      throw new ConnectionError(
        "doc.haus did not confirm this save's editing session. Check the connection and try again.",
        true,
      )
    return value
  }

  async heartbeat(base: string, generation: string): Promise<WordConnectionState> {
    return this.connection(await this.post("heartbeat", { base, generation }))
  }
  async latest(base: string): Promise<string> {
    const value = await this.post("latest", { base, input: { origin: this.link.origin } }) as { uri?: unknown }
    try {
      const uri = new URL(typeof value?.uri === "string" ? value.uri : "")
      const source = new URL(uri.searchParams.get("source") ?? "")
      if (uri.protocol !== "doc.haus:" || uri.host !== "word" || uri.pathname !== "/open" || uri.username || uri.password || uri.hash || [...uri.searchParams].length !== 1 || source.origin !== this.receiver || source.username || source.password || source.search || source.hash || !/^\/word\/open\/[A-Za-z0-9_-]{43}$/.test(source.pathname)) throw new Error()
      return uri.href
    } catch {
      throw new ConnectionError("The latest copy could not be opened. Open this matter in doc.haus and choose Open in Word.")
    }
  }
  async latestContent(base: string): Promise<LatestWordContent> {
    const value = await this.post("latest", { base, input: { origin: this.link.origin, mode: "inline" } }) as { version?: { id?: unknown; number?: unknown }; bytes?: unknown; sha256?: unknown }
    if (!value?.version || typeof value.version.id !== "string" || !value.version.id || !Number.isSafeInteger(value.version.number) || Number(value.version.number) < 1 ||
      typeof value.bytes !== "string" || !value.bytes.length || value.bytes.length > Math.ceil(4 * 1024 * 1024 / 3) * 4 ||
      !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value.bytes) || typeof value.sha256 !== "string" || !/^[a-f0-9]{64}$/.test(value.sha256))
      throw new ConnectionError("The latest document could not be verified. Your Word copy is unchanged.")
    const bytes = Uint8Array.from(atob(value.bytes), character => character.charCodeAt(0))
    if (bytes.length > 4 * 1024 * 1024) throw new ConnectionError("Open the latest copy separately; it is too large to refresh here.")
    const hash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)), byte => byte.toString(16).padStart(2, "0")).join("")
    if (hash !== value.sha256) throw new ConnectionError("The latest document could not be verified. Your Word copy is unchanged.")
    return { version: { id: value.version.id, number: Number(value.version.number) }, bytes }
  }
  async done(base: string, generation: string): Promise<WordConnectionState> {
    return this.connection(await this.post("done", { base, generation }))
  }
}
