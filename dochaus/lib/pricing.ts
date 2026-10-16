import { readFileSync } from "node:fs"

export type Rate = {
  input_per_mtok: number
  output_per_mtok: number
  cache_read_per_mtok?: number
}

export type RateOverrides = Record<string, Partial<Rate>>

type CatalogueCost = {
  input?: number
  output?: number
  cache_read?: number
  tiers?: Array<CatalogueCost & { tier: { type: string; size: number } }>
  context_over_200k?: CatalogueCost
}

type CatalogueModel = { cost?: CatalogueCost }
type Catalogue = Record<string, { models: Record<string, CatalogueModel> }>

// Read the native raw snapshot, not model.list: OpenCode normalizes missing
// prices to zero for execution. Raw data preserves unknown prices separately
// from explicitly published zero token rates (for example subscription plans).
const catalogue = loadRateCatalogue(new URL("../../packages/core/src/models-dev/snapshot.txt", import.meta.url))

export const PRICING_SOURCE = Object.keys(catalogue).length > 0
  ? "OpenCode bundled model catalogue"
  : "OpenCode bundled model catalogue unavailable"
export const PRICING_REFERENCE_URL = "https://models.opencode.ai"

export function loadRateCatalogue(file: string | URL): Catalogue {
  try {
    const parsed: unknown = JSON.parse(readFileSync(file, "utf8"))
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Catalogue : {}
  } catch {
    return {}
  }
}

export function getRate(providerModel: string, overrides?: RateOverrides, estimatedInputTokens = 0) {
  const slash = providerModel.indexOf("/")
  const model = slash > 0 ? catalogue[providerModel.slice(0, slash)]?.models?.[providerModel.slice(slash + 1)] : undefined
  return resolveRate(model, overrides?.[providerModel], estimatedInputTokens)
}

export function resolveRate(model: CatalogueModel | undefined, override?: Partial<Rate>, estimatedInputTokens = 0): Rate | undefined {
  const base = model?.cost
  const tiers = [
    ...(base?.tiers ?? []),
    ...(base?.context_over_200k ? [{ ...base.context_over_200k, tier: { type: "context", size: 200_000 } }] : []),
  ]
  // Match native SessionUsage: select the highest context threshold exceeded.
  const cost = tiers
    .filter((tier) => tier.tier.type === "context" && estimatedInputTokens > tier.tier.size)
    .sort((a, b) => b.tier.size - a.tier.size)[0] ?? base
  const input = validPrice(override?.input_per_mtok) ? override.input_per_mtok : cost?.input
  const output = validPrice(override?.output_per_mtok) ? override.output_per_mtok : cost?.output
  if (!validPrice(input) || !validPrice(output)) return undefined
  const cacheRead = validPrice(override?.cache_read_per_mtok) ? override.cache_read_per_mtok : cost?.cache_read
  return {
    input_per_mtok: input,
    output_per_mtok: output,
    ...(validPrice(cacheRead) ? { cache_read_per_mtok: cacheRead } : {}),
  }
}

export function computeCost(estimatedInputTokens: number, rate: Rate) {
  return (estimatedInputTokens / 1_000_000) * rate.input_per_mtok
}

function validPrice(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
}
