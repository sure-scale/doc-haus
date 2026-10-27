import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { httpsOrigin } from "./client"

export const WORD_BOOTSTRAP_ORIGIN = "https://doc.haus"
export const WORD_BOOTSTRAP_PATH = "/word-addin/bootstrap"
export const WORD_HEALTH_PATH = "/word-addin/health"
export const WORD_LAUNCH_URL = "doc.haus://launch"

export function wordPaneUrl(origin: string) {
  return `${httpsOrigin(origin)}/word-addin/pane.html?version=${WORD_ADDIN_VERSION}`
}

/** Only the configured origin enters this public URL; document links stay local. */
export function wordBootstrapUrl(origin: string) {
  const fragment = new URLSearchParams({ instance: httpsOrigin(origin) })
  return `${WORD_BOOTSTRAP_ORIGIN}${WORD_BOOTSTRAP_PATH}?version=${WORD_ADDIN_VERSION}#${fragment}`
}

export function loopbackWordOrigin(origin: string) {
  return ["localhost", "127.0.0.1", "[::1]"].includes(new URL(httpsOrigin(origin)).hostname)
}

export function readBootstrapOrigin(value: string): string | null {
  try {
    const url = new URL(value)
    const fragment = new URLSearchParams(url.hash.slice(1))
    if (url.origin !== WORD_BOOTSTRAP_ORIGIN || url.username || url.password ||
      ![WORD_BOOTSTRAP_PATH, `${WORD_BOOTSTRAP_PATH}.html`].includes(url.pathname) ||
      // Office may append its own opaque host metadata. Do not read or forward
      // that metadata; only our single version and instance select behavior.
      url.searchParams.getAll("version").length !== 1 || url.searchParams.get("version") !== WORD_ADDIN_VERSION ||
      fragment.size !== 1 || !fragment.has("instance")) return null
    return httpsOrigin(fragment.get("instance")!)
  } catch { return null }
}
