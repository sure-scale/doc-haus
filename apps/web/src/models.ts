// Model choices follow the user's saved or explicit provider selection.
export type ModelOption = { value: string; label: string; providerId: string; modelId: string }
export const providerOf = (value: string) => value.includes("/") ? value.slice(0, value.indexOf("/")) : ""
export function defaultProvider(models: ModelOption[]) {
  const ids = [...new Set(models.map(model => model.providerId))]
  return ids.length === 1 ? ids[0]! : ""
}
export function pickForProvider(models: ModelOption[], providerId: string) {
  const available = models.filter(model => model.providerId === providerId)
  const only = available.length === 1 ? available[0]!.value : ""
  return { primary: only, fast: only }
}
// A paid verification is explicit. Prefer the selected model; otherwise the
// provider's first available model is shown on the test action before it runs.
export function probeModel(_providerId: string, modelIds: string[]) { return modelIds[0] }
