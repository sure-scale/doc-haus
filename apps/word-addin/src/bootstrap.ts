import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { WORD_HEALTH_PATH, WORD_LAUNCH_URL, loopbackWordOrigin, readBootstrapOrigin, wordPaneUrl } from "./bootstrap-contract"

type Request = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>
type BootstrapOffice = { onReady(): Promise<unknown> }

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

async function initializedBeforeDeadline(initialized: Promise<boolean>, timeoutMs: number) {
  let timeout: ReturnType<typeof setTimeout> | undefined
  try {
    return await Promise.race([initialized, new Promise<false>(resolve => { timeout = setTimeout(() => resolve(false), timeoutMs) })])
  } finally { clearTimeout(timeout) }
}

/** This probe carries no document settings, capability, cookies or credentials. */
export async function probeWordPane(origin: string, request: Request = fetch, timeoutMs = 3000): Promise<boolean> {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), timeoutMs)
  try {
    const response = await request(`${origin}${WORD_HEALTH_PATH}`, {
      credentials: "omit", redirect: "error", referrerPolicy: "no-referrer", cache: "no-store", signal: controller.signal,
    })
    if (!response.ok) { void response.body?.cancel().catch(() => {}); return false }
    const health: unknown = await response.json()
    return Boolean(health && typeof health === "object" &&
      (health as { ready?: unknown }).ready === true && (health as { version?: unknown }).version === WORD_ADDIN_VERSION)
  } catch { return false }
  finally { clearTimeout(timeout) }
}

/** Importing this module never starts Office or reads the open document. */
export async function showBootstrap(document: Document, href: string, navigate: (url: string) => void,
  request?: Request, initialization: Promise<unknown> = Promise.resolve(), initializationTimeoutMs = 5000) {
  // Observe rejection immediately; a CDN failure must not become an unhandled
  // promise while the independent instance probe is still pending.
  const initialized = initialization.then(() => true, () => false)
  const message = document.getElementById("message")!
  const actions = document.getElementById("actions")!
  const open = document.getElementById("open") as HTMLAnchorElement
  const retry = document.getElementById("retry") as HTMLButtonElement
  const origin = readBootstrapOrigin(href)
  if (!origin) {
    message.textContent = "This Word connection needs updating. Open Settings → Add-ins in doc.haus and repair Word setup."
    actions.hidden = true
    return false
  }
  const local = loopbackWordOrigin(origin)
  open.hidden = !local
  if (local) {
    open.href = WORD_LAUNCH_URL
    open.addEventListener("click", () => {
      message.textContent = "Once doc.haus is open, choose Try again. Your Word document stays open."
    })
  }
  let checking: Promise<boolean> | undefined
  let entered = false
  const check = (): Promise<boolean> => {
    if (entered) return Promise.resolve(true)
    if (checking) return checking
    retry.disabled = true
    message.textContent = "Checking the connection to doc.haus…"
    checking = (async () => {
      if (await probeWordPane(origin, request)) {
        if (await initializedBeforeDeadline(initialized, initializationTimeoutMs)) {
          entered = true
          navigate(wordPaneUrl(origin))
          return true
        }
        message.textContent = "Word could not open the connection. Choose Try again, or check your internet connection."
      } else {
        message.textContent = local
          ? "Open doc.haus, then choose Try again. Your Word document stays open."
          : "Your doc.haus instance is not reachable. Open it in your browser or ask your administrator to check it, then choose Try again."
      }
      // Keep the useful holding page and launch action available throughout
      // startup. A retry never navigates to an unreachable instance page.
      actions.hidden = false
      return false
    })().finally(() => { retry.disabled = false; checking = undefined })
    return checking
  }
  retry.addEventListener("click", () => { void check() })
  return check()
}
