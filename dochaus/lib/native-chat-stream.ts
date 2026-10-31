import type { V2Event } from "@opencode/client"

export type NativeChatStreamOptions<Snapshot> = {
  connect: (signal: AbortSignal) => AsyncIterable<V2Event>
  signal: AbortSignal
  directory?: string
  acceptsSession: (sessionID: string) => boolean
  onEvent?: (event: V2Event) => void
  needsRefresh?: (event: V2Event) => boolean
  read: (sessionID: string, signal: AbortSignal) => Promise<Snapshot>
  onSnapshot: (sessionID: string, snapshot: Snapshot) => void
  onReadError?: (sessionID: string, error: unknown) => void
  onReconnect?: () => void
  onDisconnect?: () => void
  refreshDelayMs?: number
  reconnectDelayMs?: number
}

/** Native events are global. Never read a transcript before checking ownership. */
export function nativeEventSession(event: V2Event): string | undefined {
  const data = event.data
  if (!data || typeof data !== "object") return
  if ("sessionID" in data && typeof data.sessionID === "string") return data.sessionID
  if (event.type === "form.created") return event.data.form.sessionID
}

export function nativeChatChanged(event: V2Event): boolean {
  return event.type.startsWith("session.text.") || event.type.startsWith("session.reasoning.") ||
    event.type.startsWith("session.tool.") || event.type === "session.step.ended" || event.type === "session.step.failed"
}

/**
 * One subscription lifecycle for browser and native wrappers. The engine remains
 * authoritative; adapters hydrate their existing display projection, not a new
 * message store. A Word adapter can hydrate only independently reviewed output.
 */
export async function subscribeNativeChatStream<Snapshot>(options: NativeChatStreamOptions<Snapshot>): Promise<void> {
  const { signal } = options
  const retired = new Set<string>()
  type Refresh = { timer?: ReturnType<typeof setTimeout>; reading?: AbortController; dirty: boolean }
  const refreshes = new Map<string, Refresh>()
  const accepts = (id: string) => !signal.aborted && !retired.has(id) && options.acceptsSession(id)
  const clear = (id: string) => {
    const refresh = refreshes.get(id)
    clearTimeout(refresh?.timer)
    refresh?.reading?.abort()
    refreshes.delete(id)
  }
  const cleanup = () => { for (const id of refreshes.keys()) clear(id) }
  signal.addEventListener("abort", cleanup, { once: true })
  const queue = (id: string) => {
    if (!accepts(id)) return
    let refresh = refreshes.get(id)
    if (!refresh) { refresh = { dirty: false }; refreshes.set(id, refresh) }
    if (refresh.reading) { refresh.dirty = true; return }
    if (refresh.timer) return
    const current = refresh
    current.timer = setTimeout(() => {
      current.timer = undefined
      if (!accepts(id)) { clear(id); return }
      const reading = new AbortController()
      current.reading = reading
      void options.read(id, reading.signal).then(snapshot => {
        if (accepts(id) && !reading.signal.aborted) options.onSnapshot(id, snapshot)
      }).catch(error => {
        if (accepts(id) && !reading.signal.aborted) options.onReadError?.(id, error)
      }).finally(() => {
        if (refreshes.get(id) !== current) return
        current.reading = undefined
        if (current.dirty && accepts(id)) { current.dirty = false; queue(id) }
        else refreshes.delete(id)
      })
    }, options.refreshDelayMs ?? 100)
  }
  let connected = false
  try {
    while (!signal.aborted) {
      let began = false
      try {
        // The signal must reach fetch, otherwise idle SSE sockets survive remounts.
        for await (const event of options.connect(signal)) {
          if (signal.aborted) return
          if (!began) {
            began = true
            if (connected) options.onReconnect?.()
            connected = true
          }
          if (event.location && event.location.directory !== options.directory) continue
          const id = nativeEventSession(event)
          if (!id || !accepts(id)) continue
          if (event.type === "session.deleted") { retired.add(id); clear(id) }
          options.onEvent?.(event)
          if ((options.needsRefresh ?? nativeChatChanged)(event)) queue(id)
        }
      } catch {
        if (signal.aborted) return
      }
      if (signal.aborted) return
      options.onDisconnect?.()
      await new Promise<void>(resolve => {
        const done = () => { clearTimeout(timer); signal.removeEventListener("abort", done); resolve() }
        const timer = setTimeout(done, options.reconnectDelayMs ?? 1_000)
        signal.addEventListener("abort", done, { once: true })
      })
    }
  } finally {
    cleanup()
    signal.removeEventListener("abort", cleanup)
  }
}
