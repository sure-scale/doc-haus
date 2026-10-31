import { expect, test } from "bun:test"
import { watchWordChatSelection, type WordChatSelectionOffice } from "./word-chat-selection"

function fixture() {
  type Handler = (event: unknown) => void
  const existing: Handler = () => {}
  const listeners = new Set([existing])
  const added: { event: string; handler: Handler }[] = []
  const removed: { event: string; handler: Handler }[] = []
  let callback!: (result: { status: string }) => void
  const office: WordChatSelectionOffice = {
    EventType: { DocumentSelectionChanged: "native-selection-event" },
    context: { document: {
      addHandlerAsync(event, handler, result) {
        added.push({ event, handler })
        callback = result
      },
      removeHandlerAsync(event, { handler }, result) {
        removed.push({ event, handler })
        listeners.delete(handler)
        result({ status: "succeeded" })
      },
    } },
  }
  return {
    office, added, removed, listeners, existing,
    finish(status = "succeeded") {
      if (status === "succeeded") listeners.add(added[0]!.handler)
      callback({ status })
    },
    emit() { for (const handler of listeners) handler({ type: "native-selection-event" }) },
  }
}

test("native selection events notify until disposal, which removes only this observer", () => {
  const f = fixture()
  let changes = 0
  const dispose = watchWordChatSelection(() => { changes++ }, f.office)
  expect(f.added).toHaveLength(1)
  expect(f.added[0]!.event).toBe("native-selection-event")
  f.finish()
  f.emit()
  f.emit()
  expect(changes).toBe(2)
  dispose()
  dispose()
  expect(f.removed).toEqual([{ event: "native-selection-event", handler: f.added[0]!.handler }])
  expect(f.listeners).toEqual(new Set([f.existing]))
  // Even an event already queued by Office becomes inert after disposal.
  f.added[0]!.handler({})
  f.emit()
  expect(changes).toBe(2)
})

test("disposal before asynchronous registration completes removes its late handler and ignores queued events", () => {
  const f = fixture()
  let changes = 0
  const dispose = watchWordChatSelection(() => { changes++ }, f.office)
  dispose()
  expect(f.removed).toEqual([])
  f.added[0]!.handler({})
  f.finish()
  f.emit()
  dispose()
  expect(changes).toBe(0)
  expect(f.removed).toHaveLength(1)
  expect(f.listeners).toEqual(new Set([f.existing]))
})

test("failed registration and unsupported surfaces are safe without polling or removal of unrelated listeners", () => {
  const f = fixture()
  let changes = 0
  const dispose = watchWordChatSelection(() => { changes++ }, f.office)
  f.finish("failed")
  f.emit()
  dispose()
  expect(changes).toBe(0)
  expect(f.removed).toEqual([])
  for (const unsupported of [{}, { context: {} }, { context: { document: {} } }])
    expect(() => watchWordChatSelection(() => { changes++ }, unsupported)()).not.toThrow()
  let registrations = 0
  watchWordChatSelection(() => {}, { context: { document: { addHandlerAsync() { registrations++ } } } })()
  expect(registrations).toBe(0) // Do not create an observer that cannot be removed.
})

test("the documented string event works without the enum and synchronous host callbacks are supported", () => {
  let eventName = ""
  let handler!: (event: unknown) => void
  let changes = 0
  let removals = 0
  const dispose = watchWordChatSelection(() => { changes++ }, { context: { document: {
    addHandlerAsync(event, value, callback) {
      eventName = event
      handler = value
      callback({ status: "succeeded" })
    },
    removeHandlerAsync(event, options, callback) {
      expect(event).toBe(eventName)
      expect(options.handler).toBe(handler)
      removals++
      callback({ status: "succeeded" })
    },
  } } })
  expect(eventName).toBe("documentSelectionChanged")
  handler({})
  expect(changes).toBe(1)
  dispose()
  expect(removals).toBe(1)
})

test("host registration/removal exceptions do not escape teardown or reactivate disposed observers", () => {
  expect(() => watchWordChatSelection(() => {}, { context: { document: {
    addHandlerAsync() { throw new Error("Unsupported native event") },
    removeHandlerAsync() { throw new Error("Must not remove unrelated observers") },
  } } })()).not.toThrow()
  let handler!: (event: unknown) => void
  let changes = 0
  const dispose = watchWordChatSelection(() => { changes++ }, { context: { document: {
    addHandlerAsync(_event, value, callback) { handler = value; callback({ status: "succeeded" }) },
    removeHandlerAsync() { throw new Error("Native document has closed") },
  } } })
  expect(() => dispose()).not.toThrow()
  handler({})
  expect(changes).toBe(0)
})
