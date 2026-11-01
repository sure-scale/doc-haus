import { UPKEEP_AGENT, usageAgentKey } from "./usage-agents"

// Usage entries for one native session, priced as the provider reported it.
// Each model step carries its agent, model, tokens and cost; a session's cost
// beyond its messages is title and compaction upkeep. Ingest reads live
// sessions with this; the engine keeps a deleted session's entries with it.
export type UsageTokens = { input: number; output: number; reasoning: number; cache: { read: number; write: number } }
export type UsageEntry = { time: number; cost: number; session?: string; agent: string; model?: string; tokens?: UsageTokens }
export type UsageSession = {
  id: string
  parentID?: string
  title?: string
  agent?: string
  metadata?: Record<string, unknown> & { dochaus_role?: unknown }
  cost: number
  tokens: { input: number; output: number; reasoning: number; cache?: { read: number; write: number } }
  time: { created: number; updated: number }
}
export type UsageMessage = { info: { cost?: number; agent?: string; providerID?: string; modelID?: string; tokens?: Partial<UsageTokens>; time: { created: number } } }

export const usageTokens = (value: Partial<UsageTokens> | undefined): UsageTokens => ({
  input: value?.input || 0, output: value?.output || 0, reasoning: value?.reasoning || 0,
  cache: { read: value?.cache?.read || 0, write: value?.cache?.write || 0 },
})

// The whole session cost under the session's own agent: a session without cost,
// or one whose messages could not be read.
export function wholeSessionUsage(session: UsageSession): UsageEntry {
  return { time: session.time.created, cost: session.cost || 0, session: session.id, agent: usageAgentKey(session.agent, session), tokens: usageTokens(session.tokens) }
}

// Callers read messages only for a session with cost.
export function sessionUsageEntries(session: UsageSession, messages: readonly UsageMessage[]): UsageEntry[] {
  if (!session.cost) return [wholeSessionUsage(session)]
  const entries = messages.flatMap(({ info }): UsageEntry[] => info.cost ? [{
    time: info.time.created, cost: info.cost, session: session.id,
    agent: usageAgentKey(info.agent ?? session.agent, session),
    ...(info.providerID && info.modelID ? { model: `${info.providerID}/${info.modelID}` } : {}),
    tokens: usageTokens(info.tokens),
  }] : [])
  const rest = session.cost - entries.reduce((sum, entry) => sum + entry.cost, 0)
  // Float noise only; a title can cost less than a hundredth of a cent.
  if (rest > 1e-9) entries.push({ time: session.time.updated, cost: rest, session: session.id, agent: UPKEEP_AGENT })
  return entries
}
