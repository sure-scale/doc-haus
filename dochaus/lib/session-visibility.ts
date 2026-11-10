// Session identity comes from engine metadata, never words in a human title.
// Old sessions without a role remain visible rather than guessing ownership.
export function isUserConversation(session: { parentID?: string; metadata?: unknown }) {
  if (session.parentID) return false
  const metadata = session.metadata
  const role = typeof metadata === "object" && metadata !== null && "dochaus_role" in metadata ? metadata.dochaus_role : undefined
  return role === undefined || role === "lean-drafting-owner"
}
