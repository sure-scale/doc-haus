import { projectEvent } from "../session-view"
import type { Event } from "../session-view"
import { location, type Client } from "./client"
import { getMessages } from "./sessions"
import { providerHealth } from "./provider-health"

// Subscribe to the server event stream and invoke onEvent for each event.
// Caller passes an AbortSignal to stop. Errors after abort are swallowed.
//
// The stream is a single long-lived SSE connection. A network blip or a long
// idle gap (a heavy model thinking phase emits no parts for a while) can end it,
// and without reconnecting the live view would silently stop updating — missing
// the rest of the turn and its session.idle, so the turn appears stuck until a
// manual reload. So loop: resubscribe whenever the stream ends, and fire
// onReconnect on each re-subscribe so the caller can resync the state it missed.
export async function subscribeEvents(
  client: Client,
  onEvent: (e: Event) => void,
  signal: AbortSignal,
  options: {
    acceptsSession: (sessionID: string) => boolean
    onReconnect?: () => void
    onFormsChanged?: () => void
  },
) {
  let first = true
  const refreshTimers = new Map<string, ReturnType<typeof setTimeout>>()
  const retired = new Set<string>()
  const accepts = (sessionID: string) => !signal.aborted && !retired.has(sessionID) && options.acceptsSession(sessionID)
  const clearRefreshes = () => {
    for (const timer of refreshTimers.values()) clearTimeout(timer)
    refreshTimers.clear()
  }
  signal.addEventListener("abort", clearRefreshes, { once: true })
  const refresh = (sessionID: string) => {
    if (refreshTimers.has(sessionID)) return
    refreshTimers.set(
      sessionID,
      setTimeout(() => {
        refreshTimers.delete(sessionID)
        if (!accepts(sessionID)) return
        getMessages(client, sessionID)
          .then((messages) => {
            if (!accepts(sessionID)) return
            for (const message of messages) {
              if (message.info.role !== "assistant") continue
              for (const part of message.parts) onEvent({ type: "message.part.updated", properties: { part } })
            }
          })
          .catch(() => {})
      }, 100),
    )
  }
  try {
    while (!signal.aborted) {
      try {
        // The signal must reach the fetch itself: without it, abort only flips the
        // loop flag and the idle SSE socket stays open until the next event — each
        // panel remount then leaks a connection until the browser's per-origin cap
        // starves every other request to the engine.
        const res = client.event.subscribe({ signal })
        if (!first) options.onReconnect?.()
        first = false
        for await (const event of res) {
          if (signal.aborted) return
          // Native /api/event is global; the matter header does not scope it.
          // Ownership must be checked before scheduling a transcript read.
          const scope = location(client)
          if (event.location && event.location.directory !== scope?.directory) continue
          const sessionID =
            "sessionID" in event.data
              ? event.data.sessionID
              : event.type === "form.created"
                ? event.data.form.sessionID
                : undefined
          if (typeof sessionID !== "string" || !accepts(sessionID)) continue
          if (event.type === "session.deleted") {
            retired.add(sessionID)
            const timer = refreshTimers.get(sessionID)
            if (timer) clearTimeout(timer)
            refreshTimers.delete(sessionID)
          }
          if (event.type === "session.step.failed" && ["provider.auth", "provider.no-route"].includes(event.data.error.type)) {
            const observedVersion = providerHealth.version()
            const failure = event.data.error
            void client.session.get({ sessionID }, { signal: AbortSignal.any([signal, AbortSignal.timeout(5000)]) }).then(info => {
              if (!signal.aborted && info.model) providerHealth.rejected(info.model, failure, observedVersion)
            }).catch(() => { /* The turn still shows its native failure if model lookup is unavailable. */ })
          }
          if (event.type.startsWith("form.")) options.onFormsChanged?.()
          for (const projected of projectEvent(event)) onEvent(projected)
          if (
            (event.type.startsWith("session.text.") ||
              event.type.startsWith("session.reasoning.") ||
              event.type.startsWith("session.tool.") ||
              event.type === "session.step.ended" ||
              event.type === "session.step.failed") &&
            "sessionID" in event.data &&
            typeof event.data.sessionID === "string"
          )
            refresh(event.data.sessionID)
        }
      } catch {
        if (signal.aborted) return
      }
      if (signal.aborted) return
      await new Promise((resolve) => setTimeout(resolve, 1000))
    }
  } finally {
    clearRefreshes()
    signal.removeEventListener("abort", clearRefreshes)
  }
}
