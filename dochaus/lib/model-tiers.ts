// Per-agent model tier + thinking budget — single source of truth.
//
// A *tier* names a capability floor, not a model: "smart" is the judgment-grade
// model, "fast" is the throughput model, "fastest" is the trivial-op model. The
// tier resolves to a concrete model handle from the LIVE engine config at
// runtime, so tiering FOLLOWS the user's smart/fast model selection rather than
// pinning a fixed model. The mechanism is N-tier-capable: add a tier to
// resolveTierModel + TIER_THINKING and map agents to it; nothing else hardcodes
// the count.
//
// Two knobs move on the SAME axis — cognitive load:
//   - WHICH model runs the agent, applied by the native engine config projection,
//     which sets each tiered agent's `model` to the live resolution below. The
//     engine bakes that into the agent definition before any session runs, so a
//     subagent inherits the tier's model the same as if it were in frontmatter —
//     but it tracks config.model / config.small_model instead of a fixed string.
//   - HOW HARD it thinks (TIER_THINKING), applied dynamically by the legal
//     plugin's chat.params hook as a gemini `thinkingLevel`.
//
// Tiering is keyed by the agent's role, never by the particular findings a run
// happens to surface.

export type Tier = "smart" | "fast" | "fastest"
export type ThinkingLevel = "minimal" | "low" | "medium" | "high"

// An unconfigured install has no implicit cloud generation target. The engine
// leaves tiered agents unpinned until the user selects a primary or fast model.
export function resolveTierModel(tier: Tier, config: { model?: string; small_model?: string }): string | undefined {
  return tier === "smart" ? config.model : config.small_model ?? config.model
}

// Gemini 3 thinkingLevel per tier. "minimal" is the policy floor; the request
// adapter raises it to the lowest level a selected model actually supports.
// Mechanical fill never needs to reason, so the fast default stays at that
// floor while judgment work thinks "high".
export const TIER_THINKING: Record<Tier, ThinkingLevel> = {
  smart: "high",
  fast: "minimal",
  fastest: "minimal",
}

// Role → tier. Agents not listed here are untouched and inherit the user's
// default model and thinking behavior. The orchestrator is pinned to "fast"
// because its job is tool-calling and sequencing — exactly what the throughput
// model is optimized for.
export const AGENT_TIER: Record<string, Tier> = {
  draft: "fast",
  // The drafting narrator restates progress in plain language; it makes no legal judgment.
  narrator: "fastest",
}

export function tierOf(agent: string): Tier {
  return AGENT_TIER[agent] ?? "fast"
}

export function modelOf(agent: string, config: { model?: string; small_model?: string }): string | undefined {
  return resolveTierModel(tierOf(agent), config)
}

export function thinkingOf(agent: string): ThinkingLevel {
  return TIER_THINKING[tierOf(agent)]
}

// Thinking level ONLY for roles the tier system explicitly classifies; undefined
// for everything else so the plugin leaves their base thinking behavior
// untouched rather than forcing a budget on a role we have not deliberately tiered.
export function assignedThinkingOf(agent: string): ThinkingLevel | undefined {
  const tier = AGENT_TIER[agent]
  return tier ? TIER_THINKING[tier] : undefined
}
