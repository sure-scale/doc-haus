import { projectEvent } from "../session-view"
import type { Event } from "../session-view"
import { location, type Client } from "./client"
import { getMessages } from "./sessions"
import { providerHealth } from "./provider-health"
import { subscribeNativeChatStream } from "../../../../../dochaus/lib/native-chat-stream"

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
  return subscribeNativeChatStream({
    connect: signal => client.event.subscribe({ signal }),
    signal,
    directory: location(client)?.directory,
    acceptsSession: options.acceptsSession,
    onReconnect: options.onReconnect,
    read: (sessionID, signal) => getMessages(client, sessionID, signal),
    onSnapshot: (_sessionID, messages) => {
      for (const message of messages) {
        if (message.info.role !== "assistant") continue
        for (const part of message.parts) onEvent({ type: "message.part.updated", properties: { part } })
      }
    },
    onEvent: event => {
      if (event.type === "session.step.failed" && ["provider.auth", "provider.no-route"].includes(event.data.error.type)) {
        const observedVersion = providerHealth.version()
        const failure = event.data.error
        void client.session.get({ sessionID: event.data.sessionID }, { signal: AbortSignal.any([signal, AbortSignal.timeout(5000)]) }).then(info => {
          if (!signal.aborted && info.model) providerHealth.rejected(info.model, failure, observedVersion)
        }).catch(() => { /* The turn still shows its native failure if model lookup is unavailable. */ })
      }
      if (event.type.startsWith("form.")) options.onFormsChanged?.()
      for (const projected of projectEvent(event)) onEvent(projected)
    },
  })
}
