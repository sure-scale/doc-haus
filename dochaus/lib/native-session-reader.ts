import type { SessionDomain } from "@opencode/plugin/promise/session"
import { realpath } from "node:fs/promises"
import { projectSession, projectSessionMessages } from "./engine-v2-projection"

/** Authenticate the native global session identity before reading matter history. */
export async function nativeSessionReader(directory: string, sessions: Pick<SessionDomain, "get" | "messages">) {
  const canonicalDirectory = await realpath(directory)
  const get = async (sessionID: string) => {
    const session = projectSession(await sessions.get({ sessionID }))
    if (session.id !== sessionID || await realpath(session.directory) !== canonicalDirectory)
      throw new Error("Native session is outside the requested matter.")
    return session
  }
  return {
    get,
    read: async (sessionID: string, limit?: number) => {
      const session = await get(sessionID)
      const messages = await sessions.messages({ sessionID })
      // Retain the legacy latest-page policy before projection filters control rows.
      return { session, messages: projectSessionMessages(limit ? messages.slice(-Math.min(limit, 200)) : messages, session) }
    },
  }
}
