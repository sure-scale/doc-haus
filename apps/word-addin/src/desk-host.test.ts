import { expect, test } from "bun:test"
import { connectDeskDialog, DeskRequestError, openDeskDialog, type DeskOffice } from "./desk-host"
import { deskMessage, readDeskMessage } from "./desk-protocol"

const ORIGIN = "https://localhost:4612"

function install(expectedUrl = `${ORIGIN}/word-addin/desk.html?version=1`) {
  const original = Object.getOwnPropertyDescriptor(globalThis, "window")
  const listeners = new Map<string, Set<() => void>>()
  const parentHandlers = new Map<string, (event: { message: string; origin?: string } | { error: number }) => void>()
  let childHandler: ((event: { message: string; origin?: string }) => void) | undefined
  let closed = 0
  let supported = true
  const messages: string[] = []
  const office: DeskOffice = {
    onReady: async () => {},
    EventType: {
      DialogMessageReceived: "message",
      DialogEventReceived: "closed",
      DialogParentMessageReceived: "parent",
    },
    context: {
      requirements: { isSetSupported: () => supported },
      ui: {
        displayDialogAsync: (url, options, callback) => {
          expect(url).toBe(expectedUrl)
          expect(options.displayInIframe).toBe(false)
          callback({
            status: "succeeded",
            value: {
              addEventHandler: (type, handler) => {
                parentHandlers.set(type, handler)
              },
              messageChild: (message, options) => {
                expect(options.targetOrigin).toBe(ORIGIN)
                messages.push(message)
                childHandler?.({ message, origin: ORIGIN })
              },
              close: () => {
                closed++
              },
            },
          })
        },
        addHandlerAsync: (_event, handler, callback) => {
          childHandler = handler
          callback({ status: "succeeded", value: undefined })
        },
        messageParent: (message, options) => {
          expect(options.targetOrigin).toBe(ORIGIN)
          messages.push(message)
          parentHandlers.get("message")?.({ message, origin: ORIGIN })
        },
      },
    },
  }
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value: {
      location: { origin: ORIGIN },
      addEventListener: (event: string, listener: () => void) => {
        if (!listeners.has(event)) listeners.set(event, new Set())
        listeners.get(event)!.add(listener)
      },
      removeEventListener: (event: string, listener: () => void) => {
        listeners.get(event)?.delete(listener)
      },
    },
  })
  return {
    office,
    parentHandlers,
    messages,
    get closed() {
      return closed
    },
    set supported(value: boolean) {
      supported = value
    },
    unload: () => {
      for (const callback of [...(listeners.get("pagehide") ?? [])]) callback()
    },
    child: (event: { message: string; origin?: string }) => childHandler?.(event),
    restore: () => {
      if (original) Object.defineProperty(globalThis, "window", original)
      else delete (globalThis as { window?: unknown }).window
    },
  }
}

test("initial document view accepts only the two known modes without widening the secure URL boundary", async () => {
  for (const view of ["documents", "compare"]) {
    const url = `${ORIGIN}/word-addin/desk.html?version=1&view=${view}`
    const fixture = install(url)
    try {
      const dialog = await openDeskDialog({ url, office: fixture.office, handleRequest: async () => {} })
      dialog.close()
      for (const invalid of [
        "?version=1&version=2&view=compare",
        "?version=1&view=compare&view=documents",
        "?version=1&view=latest",
        "?version=1&view=",
        "?version=1&view=compare&token=synthetic",
        "?version=1&view=compare#fragment",
      ]) {
        await expect(openDeskDialog({ url: `/word-addin/desk.html${invalid}`, office: fixture.office, handleRequest: async () => {} })).rejects.toThrow("this doc.haus connection")
      }
    } finally { fixture.restore() }
  }
})

test("Office modeless messaging rejects foreign/missing origins, invalid actions and stale channels", async () => {
  const fixture = install()
  let calls = 0
  const dialog = await openDeskDialog({
    url: "/word-addin/desk.html?version=1",
    office: fixture.office,
    handleRequest: async () => {
      calls++
      return { sources: [] }
    },
  })
  let connection: Awaited<ReturnType<typeof connectDeskDialog>> | undefined
  try {
    connection = await connectDeskDialog({ office: fixture.office })
    expect(await connection.request({ action: "sources" })).toEqual({ sources: [] })
    expect(calls).toBe(1)
    const ready = fixture.messages.map(readDeskMessage).find((message) => message?.kind === "ready")!
    const send = fixture.parentHandlers.get("message")!
    for (const origin of [undefined, "https://foreign.example"]) {
      send({
        message: deskMessage(ready.channel, {
          kind: "request",
          id: crypto.randomUUID(),
          request: { action: "selection" },
        }),
        origin,
      })
    }
    send({
      message: deskMessage("obsolete-channel", {
        kind: "request",
        id: crypto.randomUUID(),
        request: { action: "selection" },
      }),
      origin: ORIGIN,
    })
    send({
      message: JSON.stringify({
        protocol: "dochaus.source-desk",
        version: 1,
        channel: ready.channel,
        kind: "request",
        id: crypto.randomUUID(),
        request: { action: "save" },
      }),
      origin: ORIGIN,
    })
    await Promise.resolve()
    expect(calls).toBe(1)
  } finally {
    connection?.dispose()
    dialog.close()
    fixture.restore()
  }
})

test("parent close rejects outstanding requests and child reload reconnects with a fresh channel", async () => {
  const fixture = install()
  let disconnected = 0
  let complete!: (value: unknown) => void
  const dialog = await openDeskDialog({
    url: "/word-addin/desk.html?version=1",
    office: fixture.office,
    handleRequest: () =>
      new Promise((resolve) => {
        complete = resolve
      }),
  })
  const first = await connectDeskDialog({ office: fixture.office })
  first.dispose()
  const second = await connectDeskDialog({
    office: fixture.office,
    onDisconnect: () => {
      disconnected++
    },
  })
  try {
    const pending = second.request({ action: "sources" })
    await Promise.resolve()
    fixture.unload()
    await expect(pending).rejects.toThrow("connection is closed")
    expect(fixture.closed).toBe(1)
    expect(disconnected).toBe(1)
    complete({ stale: true })
    await Promise.resolve()
    const channels = fixture.messages
      .map(readDeskMessage)
      .filter((message) => message?.kind === "ready")
      .map((message) => message!.channel)
    expect(new Set(channels).size).toBe(2)
  } finally {
    second.dispose()
    dialog.close()
    fixture.restore()
  }
})

test("only user-safe errors cross the bridge and unsupported hosts keep the pane fallback", async () => {
  const fixture = install()
  const dialog = await openDeskDialog({
    url: "/word-addin/desk.html?version=1",
    office: fixture.office,
    handleRequest: async (request) => {
      if (request.action === "sources") throw new DeskRequestError("This source changed. Choose it again.")
      throw new Error("Private host path should never be exposed")
    },
  })
  let connection: Awaited<ReturnType<typeof connectDeskDialog>> | undefined
  try {
    connection = await connectDeskDialog({ office: fixture.office })
    await expect(connection.request({ action: "sources" })).rejects.toThrow("This source changed")
    await expect(connection.request({ action: "selection" })).rejects.toThrow("could not be completed")
    expect(fixture.messages.join("\n")).not.toContain("Private host path")
    dialog.close()
    fixture.supported = false
    await expect(
      openDeskDialog({ url: "/word-addin/desk.html?version=1", office: fixture.office, handleRequest: async () => {} }),
    ).rejects.toThrow("in this pane")
    fixture.supported = true
    await expect(
      openDeskDialog({
        url: "https://foreign.example/desk.html",
        office: fixture.office,
        handleRequest: async () => {},
      }),
    ).rejects.toThrow("this doc.haus connection")
    await expect(
      openDeskDialog({ url: "/desk.html?token=never", office: fixture.office, handleRequest: async () => {} }),
    ).rejects.toThrow("this doc.haus connection")
  } finally {
    connection?.dispose()
    dialog.close()
    fixture.restore()
  }
})

test("child requests have a bounded timeout when the parent disappears without a close event", async () => {
  const fixture = install()
  const dialog = await openDeskDialog({
    url: "/word-addin/desk.html?version=1",
    office: fixture.office,
    handleRequest: () => new Promise(() => {}),
  })
  const connection = await connectDeskDialog({ office: fixture.office, timeoutMs: 10 })
  try {
    await expect(connection.request({ action: "sources" })).rejects.toThrow("did not respond")
  } finally {
    connection.dispose()
    dialog.close()
    fixture.restore()
  }
})

test("explicit child Close asks its parent to close the native dialog without invoking document actions", async () => {
  const fixture = install()
  let documentActions = 0
  let closes = 0
  const dialog = await openDeskDialog({
    url: "/word-addin/desk.html?version=1",
    office: fixture.office,
    handleRequest: async () => {
      documentActions++
    },
    onClose: () => {
      closes++
    },
  })
  const connection = await connectDeskDialog({ office: fixture.office })
  try {
    const ready = fixture.messages.map(readDeskMessage).find((message) => message?.kind === "ready")!
    const send = fixture.parentHandlers.get("message")!
    send({ message: deskMessage(ready.channel, { kind: "disconnect" }), origin: "https://foreign.example" })
    send({ message: deskMessage("obsolete-channel", { kind: "disconnect" }), origin: ORIGIN })
    expect(fixture.closed).toBe(0)
    connection.close()
    expect(fixture.closed).toBe(1)
    expect(closes).toBe(1)
    expect(documentActions).toBe(0)
    await expect(connection.request({ action: "selection" })).rejects.toThrow("connection is closed")
    connection.close()
    expect(fixture.closed).toBe(1)
  } finally {
    connection.dispose()
    dialog.close()
    fixture.restore()
  }
})
