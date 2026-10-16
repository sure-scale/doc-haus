export type ResearchConsentStatus = "pending" | "active" | "failed"

// Plugin imports and the HTTP boundary share process-local activation state.
// Missing activation must remain pending, including when native setup fails
// before the plugin can report a result. Nothing sensitive is recorded here.
const key = Symbol.for("dochaus.research-consent-status")
const shared = globalThis as typeof globalThis & { [key]?: Map<string, ResearchConsentStatus> }
const states = shared[key] ??= new Map<string, ResearchConsentStatus>()

export function researchConsentStatus(directory: string): ResearchConsentStatus {
  return states.get(directory) ?? "pending"
}

export function setResearchConsentStatus(directory: string, status: ResearchConsentStatus) {
  states.set(directory, status)
}
