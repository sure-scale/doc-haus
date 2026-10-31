import { deskMessage, isDeskRequest, isDeskDialogView, readDeskMessage, type DeskRequest, type DeskPayload } from "./desk-protocol"
import { SelectionError } from "./selection"

type Result<T> = { status: string; value: T }
type MessageEvent = { message: string; origin?: string }
type Dialog = {
  addEventHandler(event: string, handler: (event: MessageEvent | { error: number }) => void): void
  messageChild(message: string, options: { targetOrigin: string }): void
  close(): void
}

export type DeskOffice = {
  onReady(): Promise<unknown>
  EventType: { DialogMessageReceived: string; DialogEventReceived: string; DialogParentMessageReceived: string }
  context: {
    requirements: { isSetSupported(name: string, version: string): boolean }
    ui: {
      displayDialogAsync(
        url: string,
        options: { height: number; width: number; displayInIframe: boolean },
        callback: (result: Result<Dialog>) => void,
      ): void
      messageParent(message: string, options: { targetOrigin: string }): void
      addHandlerAsync(
        event: string,
        handler: (event: MessageEvent) => void,
        callback: (result: Result<void>) => void,
      ): void
    }
  }
}

export class DeskRequestError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "DeskRequestError"
  }
}

function runtime(): DeskOffice {
  const office = (window as unknown as { Office?: DeskOffice }).Office
  if (!office) throw new Error("Open matter documents or comparison from the doc.haus pane in Word.")
  return office
}

function secureOrigin(value: string): string {
  const url = new URL(value)
  if (url.protocol !== "https:" || url.username || url.password || url.origin !== value)
    throw new Error("A trusted HTTPS Word connection is required for the floating document window.")
  return url.origin
}

function supported(office: DeskOffice): boolean {
  return (
    office.context.requirements.isSetSupported("DialogApi", "1.2") &&
    office.context.requirements.isSetSupported("DialogOrigin", "1.1")
  )
}

let activeDialog = false

export async function openDeskDialog(options: {
  url: string
  handleRequest(request: DeskRequest): Promise<unknown>
  onClose?(): void
  office?: DeskOffice
  dimensions?: { height: number; width: number }
}): Promise<{ close(): void }> {
  const office = options.office ?? runtime()
  const origin = secureOrigin(window.location.origin)
  const url = new URL(options.url, origin)
  if (
    url.protocol !== "https:" ||
    url.origin !== origin ||
    url.username ||
    url.password ||
    url.hash ||
    Array.from(url.searchParams.keys()).some((key) => key !== "version" && key !== "view") ||
    url.searchParams.getAll("version").length > 1 ||
    url.searchParams.getAll("view").length > 1 ||
    (url.searchParams.has("view") && !isDeskDialogView(url.searchParams.get("view")))
  )
    throw new Error("The floating document window must use this doc.haus connection.")
  await office.onReady()
  if (!supported(office) || !office.context.ui?.displayDialogAsync)
    throw new Error("This Word version cannot open a secure floating desk. Use Open in this pane.")
  if (activeDialog) throw new Error("The floating document window is already open. Return to that window.")
  activeDialog = true
  return new Promise((resolve, reject) => {
    const opened = (result: Result<Dialog>) => {
      if (result.status !== "succeeded" || !result.value?.messageChild) {
        activeDialog = false
        result.value?.close?.()
        reject(new Error("Word could not open the floating desk. Use Open in this pane."))
        return
      }
      const dialog = result.value
      let closed = false
      let channel: string | undefined
      const seen = new Set<string>()
      const pending = new Set<string>()
      const send = (value: DeskPayload, target = channel) => {
        if (closed || !target) return
        dialog.messageChild(deskMessage(target, value), { targetOrigin: origin })
      }
      const finish = (close: boolean) => {
        if (closed) return
        try {
          send({ kind: "disconnect" })
        } catch {
          /* Parent is already unavailable. */
        }
        closed = true
        activeDialog = false
        window.removeEventListener("pagehide", leaving)
        pending.clear()
        seen.clear()
        if (close) {
          try {
            dialog.close()
          } catch {
            /* Office already closed the dialog. */
          }
        }
        options.onClose?.()
      }
      const leaving = () => finish(true)
      window.addEventListener("pagehide", leaving)
      dialog.addEventHandler(office.EventType.DialogEventReceived, () => finish(false))
      dialog.addEventHandler(office.EventType.DialogMessageReceived, (event) => {
        if (closed || !("message" in event) || event.origin !== origin) return
        const message = readDeskMessage(event.message)
        if (!message) return
        if (message.kind === "ready") {
          // A reloaded child gets a fresh channel. Late requests from its previous
          // execution cannot deliver results into this new page.
          if (channel !== message.channel) {
            channel = message.channel
            seen.clear()
            pending.clear()
          }
          try {
            send({ kind: "connected" })
          } catch {
            finish(true)
          }
          return
        }
        if (message.kind === "disconnect" && message.channel === channel) {
          finish(true)
          return
        }
        if (message.kind !== "request" || message.channel !== channel || seen.has(message.id)) return
        if (pending.size >= 4 || seen.size >= 1000) {
          try {
            send({
              kind: "response",
              id: message.id,
              ok: false,
              error: "The desk is busy. Wait for the current request, then reopen it if needed.",
            })
          } catch {
            finish(true)
          }
          return
        }
        seen.add(message.id)
        pending.add(message.id)
        const requestChannel = channel
        void Promise.resolve()
          .then(() => options.handleRequest(message.request))
          .then((value) => {
            if (!closed && channel === requestChannel) send({ kind: "response", id: message.id, ok: true, value })
          })
          .catch((error: unknown) => {
            if (closed || channel !== requestChannel) return
            const detail =
              error instanceof SelectionError || error instanceof DeskRequestError
                ? error.message.slice(0, 500)
                : "This request could not be completed. Close and reopen the document window to try again."
            try {
              send({ kind: "response", id: message.id, ok: false, error: detail })
            } catch {
              finish(true)
            }
          })
          .finally(() => {
            if (channel === requestChannel) pending.delete(message.id)
          })
      })
      resolve({ close: () => finish(true) })
    }
    try {
      const { height, width } = options.dimensions ?? { height: 75, width: 80 }
      office.context.ui.displayDialogAsync(url.href, { height, width, displayInIframe: false }, opened)
    } catch {
      activeDialog = false
      reject(new Error("Word could not open the floating desk. Use Open in this pane."))
    }
  })
}

export type DeskConnection = {
  request(request: DeskRequest): Promise<unknown>
  close(): void
  dispose(): void
}

export async function connectDeskDialog(
  options: {
    onDisconnect?(): void
    office?: DeskOffice
    timeoutMs?: number
  } = {},
): Promise<DeskConnection> {
  const office = options.office ?? runtime()
  const origin = secureOrigin(window.location.origin)
  await office.onReady()
  if (!supported(office) || !office.context.ui?.messageParent || !office.context.ui?.addHandlerAsync)
    throw new Error("This Word version cannot connect the floating desk. Use Open in this pane in Word.")
  const channel = crypto.randomUUID()
  const timeoutMs = options.timeoutMs ?? 70_000
  const pending = new Map<
    string,
    { resolve(value: unknown): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }
  >()
  let connected = false
  let disposed = false
  let handshakeTimer: ReturnType<typeof setTimeout> | undefined
  let readinessTimer: ReturnType<typeof setInterval> | undefined
  let readyResolve!: () => void
  let readyReject!: (error: Error) => void
  const ready = new Promise<void>((resolve, reject) => {
    readyResolve = resolve
    readyReject = reject
  })
  // Registration can fail before we begin awaiting readiness.
  void ready.catch(() => {})
  const dispose = (notify: boolean) => {
    if (disposed) return
    disposed = true
    connected = false
    clearTimeout(handshakeTimer)
    clearInterval(readinessTimer)
    window.removeEventListener("pagehide", leaving)
    // Office.UI exposes no removal API for this dialog event. The disposed
    // guard makes its listener inert until Office destroys the child page.
    const error = new Error("The Word pane connection is closed. Reopen matter documents or comparison from Word.")
    readyReject(error)
    for (const request of pending.values()) {
      clearTimeout(request.timer)
      request.reject(error)
    }
    pending.clear()
    if (notify) options.onDisconnect?.()
  }
  const leaving = () => dispose(false)
  const receive = (event: MessageEvent) => {
    if (disposed || event.origin !== origin) return
    const message = readDeskMessage(event.message)
    if (!message || message.channel !== channel) return
    if (message.kind === "connected") {
      connected = true
      clearTimeout(handshakeTimer)
      clearInterval(readinessTimer)
      readyResolve()
    } else if (message.kind === "disconnect") dispose(true)
    else if (message.kind === "response") {
      const request = pending.get(message.id)
      if (!request) return
      clearTimeout(request.timer)
      pending.delete(message.id)
      if (message.ok) request.resolve(message.value)
      else request.reject(new DeskRequestError(message.error))
    }
  }
  await new Promise<void>((resolve, reject) => {
    office.context.ui.addHandlerAsync(office.EventType.DialogParentMessageReceived, receive, (result) => {
      if (result.status === "succeeded") resolve()
      else reject(new Error("Word could not connect the desk. Reopen it from the Word pane."))
    })
  })
  window.addEventListener("pagehide", leaving)
  const announce = () => {
    try {
      office.context.ui.messageParent(deskMessage(channel, { kind: "ready" }), { targetOrigin: origin })
    } catch {
      dispose(true)
    }
  }
  handshakeTimer = setTimeout(() => dispose(true), timeoutMs)
  readinessTimer = setInterval(announce, 500)
  announce()
  await ready
  return {
    request: (request) => {
      if (disposed || !connected)
        return Promise.reject(new Error("The Word pane connection is closed. Reopen the desk from Word."))
      if (!isDeskRequest(request)) return Promise.reject(new Error("This document request is invalid."))
      const id = crypto.randomUUID()
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(id)
          reject(
            new Error("The Word pane did not respond. Keep the document open and reopen matter documents or comparison from Word."),
          )
        }, timeoutMs)
        pending.set(id, { resolve, reject, timer })
        try {
          office.context.ui.messageParent(deskMessage(channel, { kind: "request", id, request }), {
            targetOrigin: origin,
          })
        } catch {
          dispose(true)
        }
      })
    },
    close: () => {
      if (disposed) return
      try {
        // Only the host owns Office.Dialog.close(). window.close() is not a
        // supported substitute for closing a native Office dialog.
        office.context.ui.messageParent(deskMessage(channel, { kind: "disconnect" }), { targetOrigin: origin })
      } finally {
        dispose(false)
      }
    },
    dispose: () => dispose(false),
  }
}
