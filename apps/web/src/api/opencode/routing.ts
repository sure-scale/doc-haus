import type { Client } from "./client"
import { abortSession, createSession, deleteSession, getMessages } from "./sessions"
import { errorOf } from "../../components/chat/turn-model"

// Pick the chat's "Auto" assistant by asking the engine's small model which real
// assistant should answer. Routing runs entirely through OpenCode, so it adapts
// to whatever provider the small model lives on — Vertex, Bedrock, Azure, OpenAI,
// Ollama — and reuses the engine's own credentials and model resolution, with no
// parallel provider SDK to keep in sync. The router agent (dochaus/agent/router.md)
// has no tools and a neutral prompt; it returns one candidate name and, for the
// draft mode, the request's scope (see parseRoute). It
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
}): Promise<Route> {
  const names = input.candidates.map((c) => c.name)
  if (!names.length) throw new Error("No assistants are available.")
  const prompt = routePrompt(input)
  // Try the cheap small model first. Weak small models (Flash Lite, Haiku) can
  // return an off-list name; on that miss, escalate the same route to the primary
  // model. Primary is only paid for on a miss (and never when it equals small),
  // so the common case stays one cheap call. If neither can route, report that
  // failure instead of silently changing the user's requested operation. When a
  // router reply failed (an expired API key, an exhausted quota), report the
  // provider's own error, as an ordinary chat turn does, so the lawyer sees what
  // to fix rather than a retry hint that cannot help.
  const small = await routeOnce(input.client, input.smallModel, prompt, names)
  if (small.route) return small.route
  let error = small.error
  if (input.primaryModel && input.primaryModel !== input.smallModel) {
    const primary = await routeOnce(input.client, input.primaryModel, prompt, names)
    if (primary.route) return primary.route
    error = primary.error ?? error
  }
  throw new Error(error ?? "Auto could not choose an assistant. Try again or select an assistant manually.")
}

// One routing attempt on a "provider/model" spec: spin up the throwaway router
// session, prompt it, and return the validated route — or none on a bad spec,
// timeout, or invalid reply so the caller can escalate or report failure, with
// the reply's error text when the router turn itself failed.
async function routeOnce(client: Client, spec: string, prompt: string, names: string[]): Promise<{ route?: Route; error?: string }> {
  const slash = spec.indexOf("/")
  if (slash < 1) return {}
  const model = { providerID: spec.slice(0, slash), modelID: spec.slice(slash + 1) }
  const id = (await createSession(client, "doc.haus router").catch(() => null))?.id
  if (!id) return {}
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
    if (!res) return {}
    const replies = res.filter((message) => message.info.role === "assistant")
    const reply = replies.flatMap((message) => message.parts)
      .filter((p) => p.type === "text")
      .map((p) => p.text)
      .join("")
      .trim()
    return { route: parseRoute(reply, names), error: errorOf(replies.findLast((message) => message.info.error)?.info.error) }
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

// What a drafting request needs, as the router judged it from the message: five
// closed estimates, not verified facts about sources nobody has read yet.
export type DraftScope = {
  documents: "one" | "several"
  legal_effect: "yes" | "no"
  outside_law: "yes" | "no"
  firm_material: "yes" | "no"
  sources: "few" | "many"
}
export type Route = { name: string; scope?: DraftScope }

const SCOPE: { [K in keyof DraftScope]: readonly DraftScope[K][] } = {
  documents: ["one", "several"],
  legal_effect: ["yes", "no"],
  outside_law: ["yes", "no"],
  firm_material: ["yes", "no"],
  sources: ["few", "many"],
}

// The router replies with one JSON object (dochaus/agent/router.md): the route,
// plus the five scope fields for the draft mode. Validation is strict: an
// off-list or missing route is no route, and the scope is kept only when it has
// exactly the five fields with allowed values; anything else reads as no scope,
// which never offers a quick draft. A bare exact candidate name (no JSON) is a
// route without scope. One surrounding code fence is tolerated.
export function parseRoute(reply: string, names: string[]): Route | undefined {
  const text = reply.trim()
  if (names.includes(text)) return { name: text }
  let value: unknown
  try {
    value = JSON.parse(text.replace(/^```(?:json)?\s*([\s\S]*?)\s*```$/, "$1"))
  } catch {
    return undefined
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined
  const { route, ...fields } = value as Record<string, unknown>
  if (typeof route !== "string" || !names.includes(route)) return undefined
  return { name: route, scope: parseScope(fields) }
}

function parseScope(fields: Record<string, unknown>): DraftScope | undefined {
  const keys = Object.keys(SCOPE) as (keyof DraftScope)[]
  if (Object.keys(fields).length !== keys.length) return undefined
  for (const key of keys) if (!(SCOPE[key] as readonly unknown[]).includes(fields[key])) return undefined
  return fields as DraftScope
}

// A quick draft may be offered only when every estimate allows it: one
// document, no legal effect, no outside law, no firm template or playbook, and
// few sources. A missing or invalid scope is ineligible.
export function quickDraftEligible(scope: DraftScope | undefined) {
  return (
    scope?.documents === "one" &&
    scope.legal_effect === "no" &&
    scope.outside_law === "no" &&
    scope.firm_material === "no" &&
    scope.sources === "few"
  )
}
