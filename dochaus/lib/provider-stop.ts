/** Provider and setup failures that no retry can fix: the lawyer has to change
 * something in Settings first. The lean runner stops on them with one of these
 * messages; the web shows the same words. */
export type ProviderStopKind = "auth" | "setup" | "model"

export const PROVIDER_STOP_MESSAGES: Record<ProviderStopKind, string> = {
  auth: "The model provider refused its sign-in or account access. Reconnect the provider in Settings, then Continue resumes drafting.",
  setup: "Model setup is not finished. Complete setup in Settings, then Continue resumes drafting.",
  model: "The model this draft uses is not available from its provider. Reconnect the provider or enable the model in Settings, then Continue resumes drafting.",
}

/** The stop an engine failure calls for, read only from typed fields carried as
 * the error's `cause`: the native turn error's `type` or the engine's refusal
 * `code`. Anything else is an ordinary runtime error. */
export function providerStopKind(error: unknown): ProviderStopKind | undefined {
  const cause = error instanceof Error ? error.cause : undefined
  if (typeof cause !== "object" || cause === null) return
  const { type, code } = cause as { type?: unknown; code?: unknown }
  if (code === "setup_required") return "setup"
  if (type === "provider.auth") return "auth"
  if (type === "provider.no-route") return "model"
}
