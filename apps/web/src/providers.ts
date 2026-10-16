import { OPENCODE_URL } from "./config"
import { readBrowserJSON, writeBrowserValue } from "./browser-storage"
import type { Client } from "./api/opencode/client"
import { GATED_PROVIDERS } from "./api/opencode/provider-readiness"
export { GATED_PROVIDERS } from "./api/opencode/provider-readiness"

// Host-credential providers (Google Vertex, Vertex Anthropic, Amazon Bedrock) are
// reported "connected" by the engine on project/region presence alone — it never
// checks the gcloud ADC / AWS sign-in actually resolves (see provider.ts:494,
// autoload = Boolean(project)). So a project leaking in from the host env lists
// them as ready with the full models.dev catalog even with zero credentials, and
// they only fail at the first real call. The web gates them out of every model
// picker until a live probe (probeProvider) confirms real credentials, and
// remembers the ones that passed across reloads. Both the onboarding modal and
// full Settings share this gate so they can never disagree about what's pickable.
export const isGated = (id: string) => GATED_PROVIDERS.has(id)

// localStorage (not engine config) because it is a per-browser UI gate over the
// engine's auth, not an engine setting — the engine has no notion of "verified".
const VERIFIED_KEY = `dochaus.verifiedProviders.v2:${OPENCODE_URL}`
export function loadVerified(): Set<string> {
  const saved = readBrowserJSON(VERIFIED_KEY)
  return new Set(Array.isArray(saved) ? saved.filter((id): id is string => typeof id === "string") : [])
}
export function saveVerified(ids: Set<string>) {
  writeBrowserValue(VERIFIED_KEY, JSON.stringify([...ids]))
}

// Cached checks are scoped to this engine URL. Opening a picker filters these
// flags against its catalog and never runs a model request.
export async function revalidateVerified(_client: Client, providers: { id: string; models: Record<string, unknown> }[], connected: Set<string>): Promise<Set<string>> {
  return new Set([...loadVerified()].filter(id => connected.has(id) && providers.some(provider => provider.id === id)))
}
