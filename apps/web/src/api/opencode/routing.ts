import type { Client } from "./client"
import { abortSession, createSession, deleteSession, getMessages } from "./sessions"

// Pick the chat's "Auto" assistant by asking the engine's small model which real
// assistant should answer. Routing runs entirely through OpenCode, so it adapts
// to whatever provider the small model lives on — Vertex, Bedrock, Azure, OpenAI,
// Ollama — and reuses the engine's own credentials and model resolution, with no
// parallel provider SDK to keep in sync. The router agent (dochaus/agent/router.md)
// has no tools and a neutral prompt, so it just returns one candidate name. It
// runs on the matter's own client (passed in) — NOT a directory-less client. A
// header-less prompt routes to the server's cwd, a different directory than the
// matter, which forces the engine to cold-boot a whole second instance for cwd
// (git detection, plugin boot, provider state, models.dev fetch) on first touch,
// awaited with no timeout — the exact stall that wedged the composer on
// "Thinking...". The matter's instance is already warm (upload, agents, history
// all hit it), so routing there reuses it and never boots a second instance.
// The throwaway session is titled "doc.haus router" so the conversation rail's
// system-title filter drops it (see Sidebar), and it is deleted once read.
// smallModel is the full "provider/model" string from the engine config.
export async function routeAgent(input: {
  client: Client
  smallModel: string
  primaryModel?: string
  candidates: { name: string; description: string }[]
  history: string[]
  text: string
}): Promise<string> {
  const names = input.candidates.map((c) => c.name)
  if (!names.length) throw new Error("No assistants are available.")
  const prompt = routePrompt(input)
  // Try the cheap small model first. Weak small models (Flash Lite, Haiku) can
  // return an off-list name; on that miss, escalate the same route to the primary
  // model. Primary is only paid for on a miss (and never when it equals small),
  // so the common case stays one cheap call. If neither can route, report that
  // failure instead of silently changing the user's requested operation.
  const small = await routeOnce(input.client, input.smallModel, prompt, names)
  if (small) return small
  if (input.primaryModel && input.primaryModel !== input.smallModel) {
    const primary = await routeOnce(input.client, input.primaryModel, prompt, names)
    if (primary) return primary
  }
  throw new Error("Auto could not choose an assistant. Try again or select an assistant manually.")
}

// One routing attempt on a "provider/model" spec: spin up the throwaway router
// session, prompt it, and return the matched candidate name — or null on a bad
// spec, timeout, or off-list reply so the caller can escalate or report failure.
async function routeOnce(client: Client, spec: string, prompt: string, names: string[]): Promise<string | null> {
  const slash = spec.indexOf("/")
  if (slash < 1) return null
  const model = { providerID: spec.slice(0, slash), modelID: spec.slice(slash + 1) }
  const id = (await createSession(client, "doc.haus router").catch(() => null))?.id
  if (!id) return null
  try {
    // Routing is a best-effort pre-step that must never block the user's turn. The
    // ephemeral router turn has been seen to never complete, which would hang the
    // caller on this await (wedging the composer on "Thinking...") and — since the
    // finally would then never run — leak the throwaway session. Cap it: on overrun
    // return null and let the finally delete the session.
    const turn = (async () => {
      await client.session.switchAgent({ sessionID: id, agent: "router" })
      await client.session.switchModel({ sessionID: id, model: { providerID: model.providerID, id: model.modelID } })
      await client.session.prompt({ sessionID: id, text: prompt, delivery: "queue" })
      await client.session.wait({ sessionID: id })
      return getMessages(client, id)
    })().catch(() => null)
    // Local and open-weight models can need materially longer than a small
    // hosted model to produce their first token. Give them enough time to route
    // without silently degrading the user's request to an unrelated assistant.
    const res = await Promise.race([turn, new Promise<null>((resolve) => setTimeout(() => resolve(null), 30000))])
    if (!res) return null
    const reply = (res.filter((message) => message.info.role === "assistant").flatMap((message) => message.parts))
      .filter((p) => p.type === "text")
      .map((p) => p.text)
      .join("")
      .trim()
    // The model is told to reply with exactly one name; tolerate stray wrapping by
    // also accepting a name embedded in the reply. Off-list => null.
    return names.find((n) => reply === n) ?? names.find((n) => reply.toLowerCase().includes(n.toLowerCase())) ?? null
  } finally {
    await abortSession(client, id).catch(() => {})
    await deleteSession(client, id).catch(() => {})
  }
}

function routePrompt(input: { candidates: { name: string; description: string }[]; history: string[]; text: string }) {
  return [
    "<candidates>",
    input.candidates.map((c) => `- ${c.name}: ${c.description}`).join("\n"),
    "</candidates>",
    "<history>",
    input.history.length ? input.history.map((h) => `- ${h}`).join("\n") : "(none)",
    "</history>",
    "<message>",
    input.text,
    "</message>",
  ].join("\n")
}
