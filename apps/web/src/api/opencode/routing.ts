import type { Client } from "./client"
import { abortSession, createSession, deleteSession, getMessages } from "./sessions"
import { errorOf } from "../../components/chat/turn-model"
import { routeAgent as sharedRouteAgent, type Route } from "../../../../../dochaus/lib/chat-routing"

export { ackLine, parseRoute, quickDraftEligible, type DraftScope, type Route } from "../../../../../dochaus/lib/chat-routing"

/** Web supplies its native SDK transport; all routing decisions, retries and
 * acknowledgement validation live in the same controller Word uses. */
export function routeAgent(input: {
  client: Client
  smallModel: string
  primaryModel?: string
  candidates: { name: string; description: string }[]
  history: string[]
  text: string
  attemptMs?: number
  onAck?: (ack: string) => void
}): Promise<Route> {
  const { client, ...request } = input
  return sharedRouteAgent({ ...request, transport: {
    create: title => createSession(client, title),
    select: async (sessionID, agent, model) => {
      await client.session.switchAgent({ sessionID, agent })
      await client.session.switchModel({ sessionID, model: { providerID: model.providerID, id: model.modelID } })
    },
    prompt: async (sessionID, text) => { await client.session.prompt({ sessionID, text, delivery: "queue" }) },
    wait: async sessionID => { await client.session.wait({ sessionID }) },
    messages: async sessionID => (await getMessages(client, sessionID)).map(message => ({
      role: message.info.role,
      text: message.parts.filter(part => part.type === "text").map(part => part.text).join(""),
      error: errorOf(message.info.error),
    })),
    interrupt: async sessionID => { await abortSession(client, sessionID) },
    remove: async sessionID => { await deleteSession(client, sessionID) },
    subscribe: async (sessionID, signal, onConnected, onDelta) => {
      for await (const event of client.event.subscribe({ signal })) {
        onConnected()
        if (event.type === "session.text.delta" && event.data.sessionID === sessionID) onDelta(event.data.delta)
      }
    },
  } })
}
