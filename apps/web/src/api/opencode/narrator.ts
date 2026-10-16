import type { SessionMessageInfo } from "@opencode/client"
import { boundedRequest, clientDirectories, location, type Client } from "./client"
import { getConfig, modelSpec } from "./config"

// The drafting narrator's own session. One per drafting run, in the matter's
// directory, so its updates are durable native history a reload reads back. It
// is found again by a deterministic title naming the run's owner session, so no
// run state records it. The "doc.haus" title prefix keeps it off the
// conversation rail (see Sidebar). The narrator only reads an excerpt the
// browser sends and writes prose; it never touches the drafting run.

export const NARRATOR_AGENT = "narrator"
const NARRATION_TIMEOUT_MS = 120_000

/** One narrator update: its prose and the checkpoints its prompt covered. */
export type NarratorNote = { id: string; at: number; text: string; keys: string[] }

export function narratorTitle(ownerSessionID: string) {
  return `doc.haus narrator · ${ownerSessionID}`
}

/** The run's narrator session, oldest first if two views raced to create one. */
export async function findNarratorSession(client: Client, ownerSessionID: string) {
  const title = narratorTitle(ownerSessionID)
  const page = await boundedRequest(
    client.session.list({ limit: 20, order: "asc", search: title, directory: clientDirectories.get(client) }),
    10_000,
    "Narrator lookup",
  )
  return page.data.find((session) => session.title === title)?.id
}

export async function openNarratorSession(client: Client, ownerSessionID: string) {
  const found = await findNarratorSession(client, ownerSessionID)
  if (found) return found
  // The narrator runs on the configured small model; it never picks its own.
  const config = await boundedRequest(getConfig(), 3_000, "Model configuration")
  const model = modelSpec({ model: config.small_model ?? config.model })
  if (!model) throw new Error("No model is configured for the narrator.")
  const created = await boundedRequest(
    client.session.create({
      title: narratorTitle(ownerSessionID),
      agent: NARRATOR_AGENT,
      model: { providerID: model.providerID, id: model.modelID },
      location: location(client),
      metadata: { dochaus_role: "lean-drafting-narrator", owner_session_id: ownerSessionID },
    }),
    10_000,
    "Narrator session",
  )
  return created.id
}

/** Every update the narrator has written, each with the checkpoints its prompt
 * recorded in metadata. */
export async function readNarration(client: Client, sessionID: string) {
  const messages: SessionMessageInfo[] = []
  let cursor: string | undefined
  do {
    const page = await boundedRequest(
      client.message.list({ sessionID, limit: 200, ...(cursor ? { cursor } : { order: "asc" }) }),
      10_000,
      "Narrator history",
    )
    messages.push(...page.data)
    cursor = page.cursor.next ?? undefined
  } while (cursor)
  return narratorNotes(messages)
}

export function narratorNotes(messages: SessionMessageInfo[]): NarratorNote[] {
  let keys: string[] = []
  return messages.flatMap((message): NarratorNote[] => {
    if (message.type === "user") {
      keys = narrationKeys(message.metadata)
      return []
    }
    if (message.type !== "assistant") return []
    const text = message.content
      .flatMap((part) => (part.type === "text" ? [part.text] : []))
      .join("")
      .trim()
    if (!text) return []
    return [{ id: message.id, at: message.time.completed ?? message.time.created, text, keys }]
  })
}

/** Ask for one update and return the narrator's whole history once it settles.
 * Bounded: a stalled turn is interrupted so the next update can run. */
export async function narrate(client: Client, sessionID: string, excerpt: string, keys: string[]) {
  const turn = (async () => {
    await client.session.prompt({
      sessionID,
      text: excerpt,
      delivery: "queue",
      metadata: { dochaus_narration: { keys } },
    })
    await client.session.wait({ sessionID })
    return readNarration(client, sessionID)
  })()
  return boundedRequest(turn, NARRATION_TIMEOUT_MS, "Narration").catch(async (error: unknown) => {
    await client.session.interrupt({ sessionID }).catch(() => {})
    throw error
  })
}

function narrationKeys(metadata: Record<string, unknown> | undefined) {
  const narration = metadata?.dochaus_narration
  if (typeof narration !== "object" || narration === null || !("keys" in narration) || !Array.isArray(narration.keys))
    return []
  return narration.keys.filter((key): key is string => typeof key === "string")
}
