import { createOpencodeClient } from "@opencode-ai/sdk"
import type { Event, Part } from "@opencode-ai/sdk"
import { OPENCODE_URL } from "../config"

// One OpenCode client per matter. The matter's directory is sent as the
// x-opencode-directory header on every request, so every session and every
// tool (including search-document) is scoped to that matter's files + DB.
export function matterClient(directory: string) {
  return createOpencodeClient({ baseUrl: OPENCODE_URL, directory })
}

export type Client = ReturnType<typeof matterClient>

// A client with no matter directory. Settings act on the engine's global config
// (opencode.json) and machine-wide credential store (auth.json) rather than a
// single matter, so we deliberately omit the x-opencode-directory header.
export function settingsClient() {
  return createOpencodeClient({ baseUrl: OPENCODE_URL })
}

// All providers the engine knows from the models.dev catalog, with which ones
// are connected and the current default model per provider.
export async function listProviders(client: Client) {
  const res = await client.provider.list()
  return res.data ?? { all: [], default: {}, connected: [] }
}

// Per-provider auth methods (api key vs oauth). Empty array => credentials come
// from the environment (Vertex ADC, Bedrock/Azure SDK creds), not a key field.
export async function listAuthMethods(client: Client) {
  const res = await client.provider.auth()
  return res.data ?? {}
}

// Store an API key for a provider. Writes to the engine's auth.json (0o600).
export async function setProviderKey(client: Client, id: string, key: string) {
  return client.auth.set({ path: { id }, body: { type: "api", key } })
}

export async function getConfig(client: Client) {
  const res = await client.config.get()
  return res.data ?? {}
}

// Set the engine-wide default model, as a "providerID/modelID" string.
export async function setDefaultModel(client: Client, model: string) {
  return client.config.update({ body: { model } })
}

// Hide providers from routing entirely. A disabled provider drops out of the
// catalog, the model picker, and any agent that would route to it.
export async function setDisabledProviders(client: Client, ids: string[]) {
  return client.config.update({ body: { disabled_providers: ids } })
}

// Register a local OpenAI-compatible provider (LM Studio, Ollama, vLLM...). The
// engine loads it through @ai-sdk/openai-compatible at the given baseURL. The
// baseURL must be reachable from where `opencode serve` runs, not the browser.
export async function addLocalProvider(
  client: Client,
  input: { id: string; name: string; baseURL: string; modelID: string },
) {
  const cfg = await getConfig(client)
  return client.config.update({
    body: {
      provider: {
        ...(cfg.provider ?? {}),
        [input.id]: {
          npm: "@ai-sdk/openai-compatible",
          name: input.name,
          options: { baseURL: input.baseURL },
          models: { [input.modelID]: { name: input.modelID } },
        },
      },
    },
  })
}

// Shape returned by the search-document tool in its part metadata.citations.
export type Citation = {
  documentName: string
  docPath: string
  section: string
  excerpt: string
  charStart: number
  charEnd: number
  score: number
}

export async function listAgents(client: Client) {
  const res = await client.app.agents()
  return res.data ?? []
}

export async function createSession(client: Client, title: string) {
  const res = await client.session.create({ body: { title } })
  if (!res.data) throw new Error("Failed to create session")
  return res.data
}

// Every session the engine holds for this matter (scoped by the client's
// directory header). Subagent runs carry a parentID; top-level chats do not.
export async function listSessions(client: Client) {
  const res = await client.session.list()
  return res.data ?? []
}

// Settled messages for one session, each as { info, parts }, used to replay a
// past conversation back into the chat panel.
export async function getMessages(client: Client, sessionID: string) {
  const res = await client.session.messages({ path: { id: sessionID } })
  return res.data ?? []
}

// Fire a prompt to a named agent. Resolves when the assistant turn completes;
// live progress arrives separately through subscribeEvents.
export async function sendPrompt(client: Client, sessionID: string, agent: string, text: string) {
  return client.session.prompt({
    path: { id: sessionID },
    body: { agent, parts: [{ type: "text", text }] },
  })
}

// Subscribe to the server event stream and invoke onEvent for each event.
// Caller passes an AbortSignal to stop. Errors after abort are swallowed.
export async function subscribeEvents(client: Client, onEvent: (e: Event) => void, signal: AbortSignal) {
  const res = await client.event.subscribe()
  for await (const event of res.stream) {
    if (signal.aborted) return
    onEvent(event as Event)
  }
}

export function isToolPart(part: Part): part is Extract<Part, { type: "tool" }> {
  return part.type === "tool"
}

export function isTextPart(part: Part): part is Extract<Part, { type: "text" }> {
  return part.type === "text"
}
