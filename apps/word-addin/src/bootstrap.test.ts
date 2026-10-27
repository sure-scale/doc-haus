import { expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { Window } from "happy-dom"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { WORD_BOOTSTRAP_ORIGIN, loopbackWordOrigin, readBootstrapOrigin, wordBootstrapUrl, wordPaneUrl } from "./bootstrap-contract"
import { initializeBootstrapOffice, probeWordPane, showBootstrap } from "./bootstrap"

const origin = "https://localhost:4511"
const html = readFileSync(new URL("../bootstrap.html", import.meta.url), "utf8")
function page() {
  const window = new Window()
  window.document.body.innerHTML = html.match(/<body>([\s\S]*)<\/body>/)![1]!
  return { window, document: window.document as unknown as Document, close() {
    window.dispatchEvent(new window.Event("pagehide"))
    window.happyDOM.abort()
  } }
}

function clock() {
  const pending = new Map<number, () => void>()
  const delays: number[] = []
  let next = 0
  return {
    pending, delays,
    schedule(callback: () => void, delayMs: number) {
      const id = ++next
      delays.push(delayMs)
      pending.set(id, callback)
      return () => { pending.delete(id) }
    },
    tick() {
      const callbacks = Array.from(pending.values())
      pending.clear()
      for (const callback of callbacks) callback()
    },
  }
}

test("bootstrap accepts only one configured HTTPS origin and navigates a fixed current pane path", () => {
  for (const instance of [origin, "https://[::1]:4511", "https://word.firm.example"]) {
    const url = wordBootstrapUrl(instance)
    expect(new URL(url).origin).toBe(WORD_BOOTSTRAP_ORIGIN)
    expect(new URL(url).pathname).toBe("/word-addin/bootstrap")
    expect(new URL(url).search).toBe(`?version=${WORD_ADDIN_VERSION}`)
    expect(readBootstrapOrigin(url)).toBe(instance)
    expect(readBootstrapOrigin(url.replace("/word-addin/bootstrap?", "/word-addin/bootstrap.html?"))).toBe(instance)
    expect(wordPaneUrl(instance)).toBe(`${instance}/word-addin/pane.html?version=${WORD_ADDIN_VERSION}`)
  }
  expect(loopbackWordOrigin("https://127.0.0.1:4511")).toBe(true)
  expect(loopbackWordOrigin("https://word.firm.example")).toBe(false)
  const officeUrl = new URL(wordBootstrapUrl(origin))
  officeUrl.searchParams.append("_host_Info", "synthetic-host-metadata")
  expect(readBootstrapOrigin(officeUrl.href)).toBe(origin)
  for (const value of [
    wordBootstrapUrl(origin).replace(`version=${WORD_ADDIN_VERSION}`, "version=old"),
    `${wordBootstrapUrl(origin)}&instance=${encodeURIComponent("https://other.example")}`,
    `${wordBootstrapUrl(origin)}&token=private`,
    wordBootstrapUrl(origin).replace("#", `&version=${WORD_ADDIN_VERSION}#`),
    wordBootstrapUrl(origin).replace("https://doc.haus", "https://unrelated.example"),
    wordBootstrapUrl(origin).replace("https://doc.haus", "https://user@doc.haus"),
    wordBootstrapUrl(origin).replace("/bootstrap?", "/bootstrap/?"),
    wordBootstrapUrl(origin).replace("/bootstrap?", "/bootstrap.other?"),
    `${WORD_BOOTSTRAP_ORIGIN}/word-addin/bootstrap.html?version=${WORD_ADDIN_VERSION}#instance=${encodeURIComponent("http://localhost:4511")}`,
    `${WORD_BOOTSTRAP_ORIGIN}/word-addin/bootstrap.html?version=${WORD_ADDIN_VERSION}#instance=${encodeURIComponent("https://localhost:4511/private?token=secret")}`,
  ]) expect(readBootstrapOrigin(value)).toBeNull()
})

test("ready bootstrap probes without credentials then enters the instance pane", async () => {
  const fixture = page()
  const requests: { url: string; init?: RequestInit }[] = [], navigations: string[] = []
  try {
    expect(await showBootstrap(fixture.document, wordBootstrapUrl(origin), url => navigations.push(url), async (url, init) => {
      requests.push({ url: String(url), init })
      return Response.json({ ready: true, version: WORD_ADDIN_VERSION })
    })).toBe(true)
    expect(requests).toHaveLength(1)
    expect(requests[0]!.url).toBe(`${origin}/word-addin/health`)
    expect(requests[0]!.init).toMatchObject({ credentials: "omit", redirect: "error", referrerPolicy: "no-referrer", cache: "no-store" })
    expect(requests[0]!.init?.headers).toBeUndefined()
    expect(navigations).toEqual([wordPaneUrl(origin)])
    expect(html.match(/https:\/\/[^\"]+/g)).toEqual(["https://appsforoffice.microsoft.com/lib/1/hosted/office.js"])
  } finally { fixture.close() }
})

test("stopped local instance automatically checks again and retains a user-initiated native launch link", async () => {
  const fixture = page(), timer = clock(), navigations: string[] = []
  try {
    expect(await showBootstrap(fixture.document, wordBootstrapUrl(origin), url => navigations.push(url), async () => {
      throw new TypeError("offline")
    }, new Promise(() => {}), 5000, timer.schedule)).toBe(false)
    expect(navigations).toEqual([])
    expect(fixture.document.getElementById("actions")!.hidden).toBe(false)
    const launch = fixture.document.getElementById("open") as HTMLAnchorElement
    expect(launch.hidden).toBe(false)
    expect(launch.href).toBe("doc.haus://launch")
    expect(launch.target).toBe("_blank")
    expect(launch.relList.contains("noopener")).toBe(true)
    expect(fixture.document.getElementById("retry")).toBeNull()
    expect(timer.pending.size).toBe(1)
    expect(timer.delays).toEqual([1000])
    launch.dispatchEvent(new fixture.window.MouseEvent("click", { bubbles: true }) as unknown as Event)
    expect(fixture.document.getElementById("message")!.textContent).toBe("Open doc.haus to reconnect. Your Word document stays open.")
    expect(navigations).toEqual([])
  } finally { fixture.close() }
  expect(timer.pending.size).toBe(0)
})

test("polling stays quiet through startup, serializes probes and waits for Office before entering exactly once", async () => {
  const fixture = page(), timer = clock(), navigations: string[] = []
  let requests = 0
  let available = false
  let finishProbe!: (response: Response) => void
  let finishOffice!: () => void
  const initialization = new Promise<void>(resolve => { finishOffice = resolve })
  const request = async () => {
    requests++
    if (!available) return new Response(null, { status: 503 })
    return new Promise<Response>(resolve => { finishProbe = resolve })
  }
  const mutations: unknown[] = []
  const observer = new fixture.window.MutationObserver(records => mutations.push(...records))
  try {
    expect(await showBootstrap(fixture.document, wordBootstrapUrl(origin), url => navigations.push(url), request, initialization, 5000, timer.schedule)).toBe(false)
    observer.observe(fixture.window.document.getElementById("message")!, { childList: true, characterData: true, subtree: true })
    timer.tick()
    await Bun.sleep(0)
    expect(requests).toBe(2)
    expect(mutations).toHaveLength(0)
    expect(navigations).toEqual([])
    expect(fixture.document.getElementById("actions")!.hidden).toBe(false)
    available = true
    timer.tick()
    timer.tick()
    await Bun.sleep(0)
    expect(requests).toBe(3)
    expect(timer.pending.size).toBe(0)
    expect(fixture.document.getElementById("open")!.hidden).toBe(false)
    finishProbe(Response.json({ ready: true, version: WORD_ADDIN_VERSION }))
    await Bun.sleep(0)
    expect(navigations).toEqual([])
    finishOffice()
    await Bun.sleep(0)
    expect(navigations).toEqual([wordPaneUrl(origin)])
    expect(timer.pending.size).toBe(0)
    timer.tick()
    await Bun.sleep(0)
    expect(requests).toBe(3)
    expect(navigations).toHaveLength(1)
    expect(timer.delays).toEqual([1000, 1000])
  } finally { observer.disconnect(); fixture.close() }
})

test("pagehide cancels queued polling and an active probe without navigating or scheduling more work", async () => {
  for (const active of [false, true]) {
    const fixture = page(), timer = clock(), navigations: string[] = []
    let requests = 0
    let aborted = false
    const opening = showBootstrap(fixture.document, wordBootstrapUrl(origin), url => navigations.push(url), async (_url, init) => {
      requests++
      if (!active) return new Response(null, { status: 503 })
      return new Promise<Response>((_resolve, reject) => {
        init!.signal!.addEventListener("abort", () => {
          aborted = true
          reject(new DOMException("closed", "AbortError"))
        }, { once: true })
      })
    }, Promise.resolve(), 5000, timer.schedule)
    try {
      if (!active) {
        expect(await opening).toBe(false)
        expect(timer.pending.size).toBe(1)
      }
      fixture.window.dispatchEvent(new fixture.window.Event("pagehide"))
      expect(await opening).toBe(false)
      expect(aborted).toBe(active)
      expect(timer.pending.size).toBe(0)
      timer.tick()
      await Bun.sleep(0)
      expect(requests).toBe(1)
      expect(navigations).toHaveLength(0)
    } finally { fixture.close() }
  }
})

test("healthy instance waits for the Office handshake while stalled initialization keeps the holding page visible", async () => {
  const fixture = page(), navigations: string[] = []
  let ready!: () => void
  const initialization = initializeBootstrapOffice(fixture.document, () => ({ onReady: () => new Promise<void>(resolve => { ready = resolve }) }))
  const request = async () => Response.json({ ready: true, version: WORD_ADDIN_VERSION })
  try {
    const opening = showBootstrap(fixture.document, wordBootstrapUrl(origin), url => navigations.push(url), request, initialization)
    await Bun.sleep(0)
    expect(navigations).toEqual([])
    ready()
    expect(await opening).toBe(true)
    expect(navigations).toEqual([wordPaneUrl(origin)])
    navigations.length = 0
    expect(await showBootstrap(fixture.document, wordBootstrapUrl(origin), url => navigations.push(url), request, new Promise(() => {}), 15)).toBe(false)
    expect(navigations).toEqual([])
    expect(fixture.document.getElementById("actions")!.hidden).toBe(false)
    expect(fixture.document.getElementById("message")!.textContent).toContain("Waiting for Word")
  } finally { fixture.close() }
})

test("remote failure never claims the local launcher can start a firm server; invalid config makes no request", async () => {
  const fixture = page()
  let requests = 0
  const request = async () => { requests++; return new Response(null, { status: 503 }) }
  try {
    expect(await showBootstrap(fixture.document, wordBootstrapUrl("https://word.firm.example"), () => { throw new Error("unexpected navigation") }, request)).toBe(false)
    expect(fixture.document.getElementById("open")!.hidden).toBe(true)
    expect(fixture.document.getElementById("message")!.textContent).toContain("administrator")
    expect(requests).toBe(1)
    await showBootstrap(fixture.document, "https://doc.haus/word-addin/bootstrap.html#instance=invalid", () => {}, request)
    expect(fixture.document.getElementById("actions")!.hidden).toBe(true)
    expect(requests).toBe(1)
  } finally { fixture.close() }
})

test("unavailable, unconfirmed, old or timed-out health never reports ready", async () => {
  for (const response of [
    Response.json({ ready: true, version: "old" }), Response.json({ ready: false, version: WORD_ADDIN_VERSION }),
    Response.json({ ready: true, version: WORD_ADDIN_VERSION }, { status: 503 }), new Response("not JSON"),
  ]) expect(await probeWordPane(origin, async () => response)).toBe(false)
  let aborted = false
  expect(await probeWordPane(origin, async (_url, init) => new Promise((_resolve, reject) => {
    init!.signal!.addEventListener("abort", () => { aborted = true; reject(new DOMException("timed out", "AbortError")) }, { once: true })
  }), 15)).toBe(false)
  expect(aborted).toBe(true)
  expect(await probeWordPane(origin, async (_url, init) => new Response(new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode('{"ready":'))
      init!.signal!.addEventListener("abort", () => controller.error(new DOMException("timed out", "AbortError")), { once: true })
    },
  })), 15)).toBe(false)
})
