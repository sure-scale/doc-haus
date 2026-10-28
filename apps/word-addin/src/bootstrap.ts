import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { WORD_HEALTH_PATH, WORD_LAUNCH_URL, loopbackWordOrigin, readBootstrapOrigin, wordPaneUrl } from "./bootstrap-contract"

type Request = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>
type BootstrapOffice = { onReady(): Promise<unknown> }
type Schedule = (callback: () => void, delayMs: number) => () => void

const scheduleCheck: Schedule = (callback, delayMs) => {
  const timer = setTimeout(callback, delayMs)
  return () => clearTimeout(timer)
}

/** Office requires a readiness handshake on some hosts, even for a holding
 * page. This deliberately never obtains Office.context or document settings. */
export function initializeBootstrapOffice(document: Document, office: () => BootstrapOffice | undefined): Promise<void> {
  return new Promise((resolve, reject) => {
    let started = false
    const initialize = () => {
      const api = office()
      if (!api || started) return
      started = true
      try { void api.onReady().then(() => resolve(), reject) }
      catch { reject(new Error("Word initialization failed")) }
    }
    const script = document.getElementById("office-api")
    script?.addEventListener("load", initialize, { once: true })
    script?.addEventListener("error", () => reject(new Error("Word initialization unavailable")), { once: true })
    initialize()
  })
}

async function initializedBeforeDeadline(initialized: Promise<boolean>, timeoutMs: number, signal: AbortSignal) {
  if (signal.aborted) return false
  let timeout: ReturnType<typeof setTimeout> | undefined
  let abort: () => void = () => {}
  try {
    return await Promise.race([initialized, new Promise<false>(resolve => {
      abort = () => resolve(false)
      signal.addEventListener("abort", abort, { once: true })
      timeout = setTimeout(() => resolve(false), timeoutMs)
    })])
  } finally { clearTimeout(timeout); signal.removeEventListener("abort", abort) }
}

/** This probe carries no document settings, capability, cookies or credentials. */
export async function probeWordPane(origin: string, request: Request = fetch, timeoutMs = 3000, signal?: AbortSignal): Promise<boolean> {
  if (signal?.aborted) return false
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), timeoutMs)
  const abort = () => { clearTimeout(timeout); controller.abort() }
  signal?.addEventListener("abort", abort, { once: true })
  try {
    const response = await request(`${origin}${WORD_HEALTH_PATH}`, {
      credentials: "omit", redirect: "error", referrerPolicy: "no-referrer", cache: "no-store", signal: controller.signal,
    })
    if (!response.ok) { void response.body?.cancel().catch(() => {}); return false }
    const health: unknown = await response.json()
    return Boolean(health && typeof health === "object" &&
      (health as { ready?: unknown }).ready === true && (health as { version?: unknown }).version === WORD_ADDIN_VERSION)
  } catch { return false }
  finally { clearTimeout(timeout); signal?.removeEventListener("abort", abort) }
}

/** Importing this module never starts Office or reads the open document. */
export async function showBootstrap(document: Document, href: string, navigate: (url: string) => void,
  request?: Request, initialization: Promise<unknown> = Promise.resolve(), initializationTimeoutMs = 5000,
  schedule: Schedule = scheduleCheck) {
  // Observe rejection immediately; a CDN failure must not become an unhandled
  // promise while the independent instance probe is still pending.
  const initialized = initialization.then(() => true, () => false)
  const message = document.getElementById("message")!
  const actions = document.getElementById("actions")!
  const open = document.getElementById("open") as HTMLAnchorElement
  const origin = readBootstrapOrigin(href)
  if (!origin) {
    message.textContent = "This Word connection needs updating. Open Settings → Word add-in in doc.haus and repair Word setup."
    actions.hidden = true
    return false
  }
  const local = loopbackWordOrigin(origin)
  open.hidden = !local
  if (local) open.href = WORD_LAUNCH_URL
  const lifetime = new AbortController()
  let cancelNext: (() => void) | undefined
  const stop = () => {
    lifetime.abort()
    cancelNext?.()
    cancelNext = undefined
    document.defaultView?.removeEventListener("pagehide", stop)
  }
  document.defaultView?.addEventListener("pagehide", stop, { once: true })
  const setMessage = (text: string) => { if (message.textContent !== text) message.textContent = text }
  let checking: Promise<boolean> | undefined
  let entered = false
  const check = (): Promise<boolean> => {
    if (entered || lifetime.signal.aborted) return Promise.resolve(entered)
    if (checking) return checking
    checking = (async () => {
      const ready = await probeWordPane(origin, request, 3000, lifetime.signal)
      if (lifetime.signal.aborted) return false
      if (ready) {
        if (await initializedBeforeDeadline(initialized, initializationTimeoutMs, lifetime.signal)) {
          if (lifetime.signal.aborted) return false
          entered = true
          stop()
          navigate(wordPaneUrl(origin))
          return true
        }
        if (lifetime.signal.aborted) return false
        setMessage("Waiting for Word to connect. Check your internet connection.")
      } else {
        setMessage(local
          ? "Open doc.haus to reconnect. Your Word document stays open."
          : "Your doc.haus instance is not reachable. Open it in your browser or ask your administrator to check it. Your Word document stays open.")
      }
      // Keep the useful holding page and launch action available throughout
      // startup. Polling never navigates to an unreachable instance page.
      actions.hidden = false
      return false
    })().finally(() => {
      checking = undefined
      if (!lifetime.signal.aborted) cancelNext = schedule(() => { cancelNext = undefined; void check() }, 1000)
    })
    return checking
  }
  return check()
}
