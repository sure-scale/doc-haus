const SESSION_AGENTS = new Map<string, string>()

export function recordSessionAgent(sessionID: string, agent: string) {
  if (!SESSION_AGENTS.has(sessionID) && SESSION_AGENTS.size >= 512) {
    const oldest = SESSION_AGENTS.keys().next().value
    if (oldest !== undefined) SESSION_AGENTS.delete(oldest)
  }
  SESSION_AGENTS.set(sessionID, agent)
}

export function sessionAgent(sessionID: string) {
  return SESSION_AGENTS.get(sessionID)
}
