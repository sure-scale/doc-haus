import { projectSession, projectSessionMessages } from "./engine-v2-projection"

// Document workflows consume a stable transcript projection. All requests use
// native v2 routes and durable admission; no alternate execution loop lives here.
const adapters = new WeakSet<typeof fetch>()

export function engineV2Fetch(transport: typeof fetch = fetch): typeof fetch {
  if (adapters.has(transport)) return transport
  const request = async (input: Parameters<typeof fetch>[0], init?: RequestInit): Promise<Response> => {
    const url = new URL(input instanceof Request ? input.url : String(input))
    if (!url.pathname.startsWith("/api/")) throw new Error("Engine requests must use the v2 API.")
    const headers = new Headers(init?.headers ?? (input instanceof Request ? input.headers : undefined))
    const password = process.env.OPENCODE_SERVER_PASSWORD
    if (password && !headers.has("Authorization")) headers.set("Authorization", `Basic ${Buffer.from(`opencode:${password}`).toString("base64")}`)
    const options = { ...init, headers }
    const directory = headers.get("x-opencode-directory")
    const method = options.method ?? "GET"
    const sessionRoute = /^\/api\/session\/([^/]+)$/.exec(url.pathname)
    const messageRoute = /^\/api\/session\/([^/]+)\/message(?:\/([^/]+))?$/.exec(url.pathname)
    const promptRoute = /^\/api\/session\/([^/]+)\/prompt$/.exec(url.pathname)
    const nestedSessionRoute = /^\/api\/session\/([^/]+)\//.exec(url.pathname)
    const body: unknown = typeof options.body === "string" ? JSON.parse(options.body) : undefined
    if (url.pathname === "/api/session" && method === "POST") {
      if (!record(body) || !directory) throw new Error("Session creation requires a matter directory.")
      if (body.parentID) url.pathname = "/api/dochaus/session"
      options.body = JSON.stringify({ ...body, ...(!body.parentID ? { location: { directory } } : {}), ...(record(body.model) ? { model: nativeModel(body.model) } : {}) })
    }
    if (url.pathname === "/api/session" && method === "GET") {
      if (!directory) throw new Error("Session listing requires a matter directory.")
      url.searchParams.set("directory", directory)
      if (url.searchParams.get("roots") === "true") url.searchParams.set("parentID", "null")
      url.searchParams.delete("roots")
    }
    const allMessages = Boolean(messageRoute && !messageRoute[2] && !url.searchParams.has("limit"))
    if (messageRoute && !messageRoute[2]) {
      url.searchParams.set("order", "desc")
      if (!url.searchParams.has("limit")) url.searchParams.set("limit", "200")
    }
    // Native session routes address global IDs. The directory header selects a
    // location, but does not prove that the target session belongs to the matter.
    const targetSessionID = nestedSessionRoute?.[1] ?? (sessionRoute && method !== "GET" ? sessionRoute[1] : undefined)
    const owner = targetSessionID ? await transport(new URL(`/api/session/${targetSessionID}`, url), { headers, signal: options.signal }) : undefined
    if (owner && !owner.ok) return owner
    const ownerValue = owner ? data(await owner.json()) : undefined
    const ownerSession = owner ? projectSession(ownerValue) : undefined
    if (ownerSession && (!directory || ownerSession.directory !== directory)) throw new Error("Engine session is outside the requested matter.")
    if (promptRoute && method === "POST") {
      if (!record(body)) throw new Error("Prompt requires an object body.")
      if (!record(ownerValue) || record(body.model) && (!record(ownerValue.model) || JSON.stringify(nativeModel(body.model)) !== JSON.stringify(nativeModel(ownerValue.model))) || typeof body.agent === "string" && ownerValue.agent !== body.agent)
        return Response.json({ error: "Prompt agent/model must match its session selection." }, { status: 409 })
      const text = typeof body.text === "string" ? body.text : Array.isArray(body.parts) ? body.parts.filter(record).filter(part => part.type === "text").map(part => part.text).join("") : undefined
      if (typeof text !== "string" || Array.isArray(body.parts) && body.parts.some(part => !record(part) || part.type !== "text"))
        throw new Error("Workflow prompts require text; attachments must use the native file contract.")
      options.body = JSON.stringify({
        id: body.id ?? body.messageID, text, delivery: body.delivery ?? "queue",
        ...(body.resume !== undefined ? { resume: body.resume } : {}),
        metadata: { ...(record(body.metadata) ? body.metadata : {}), ...(record(body.tools) ? { dochaus_tools: body.tools } : {}) },
      })
    }
    const response = await transport(url, options)
    // The native interrupt answers { interrupted } without the v2 data envelope.
    if (method === "POST" && /^\/api\/session\/[^/]+\/interrupt$/.test(url.pathname)) return response
    if (promptRoute && response.status === 204) throw new Error("Engine did not acknowledge the exact durable prompt admission.")
    if (!response.ok || response.status === 204) return response
    const envelope: unknown = await response.json()
    const value = data(envelope)
    if (promptRoute) {
      const submitted: unknown = typeof options.body === "string" ? JSON.parse(options.body) : undefined
      if (!record(value) || !record(submitted) || value.id !== submitted.id || value.sessionID !== decodeURIComponent(promptRoute[1]!) || value.type !== "user" || !record(value.payload) || value.payload.text !== submitted.text || value.delivery !== submitted.delivery)
        throw new Error("Engine did not acknowledge the exact durable prompt admission.")
      return projected(response, value)
    }
    if (messageRoute) {
      const session = ownerSession!

      if (messageRoute[2]) return projected(response, projectSessionMessages([value], session)[0] ?? null)
      if (!Array.isArray(value)) throw new Error("Engine returned an invalid transcript page.")
      const messages = [...value]
      let page = envelope
      const seen = new Set<string>()
      while (allMessages && record(page) && record(page.cursor) && typeof page.cursor.next === "string") {
        const cursor = page.cursor.next
        if (seen.has(cursor)) throw new Error("Engine repeated a transcript page cursor.")
        seen.add(cursor)
        url.searchParams.delete("order")
        url.searchParams.set("cursor", cursor)
        const next = await transport(url, options)
        if (!next.ok) return next
        page = await next.json()
        const rows = data(page)
        if (!Array.isArray(rows)) throw new Error("Engine returned an invalid transcript page.")
        messages.push(...rows)
      }
      return projected(response, projectSessionMessages(messages.toReversed(), session))
    }
    if (sessionRoute && sessionRoute[1] !== "active" || url.pathname === "/api/dochaus/session" || url.pathname === "/api/session" && method === "POST") {
      const session = projectSession(value)
      if (directory && session.directory !== directory) throw new Error("Engine session is outside the requested matter.")
      return projected(response, session)
    }
    if (url.pathname === "/api/session") {
      if (!Array.isArray(value)) throw new Error("Engine returned an invalid session page.")
      const sessions = value.map(projectSession)
      let page = envelope
      const seen = new Set<string>()
      while (!url.searchParams.has("limit") && record(page) && record(page.cursor) && typeof page.cursor.next === "string") {
        const cursor = page.cursor.next
        if (seen.has(cursor)) throw new Error("Engine repeated a session page cursor.")
        seen.add(cursor)
        url.searchParams.delete("order")
        url.searchParams.set("cursor", cursor)
        const next = await transport(url, options)
        if (!next.ok) return next
        page = await next.json()
        const rows = data(page)
        if (!Array.isArray(rows)) throw new Error("Engine returned an invalid session page.")
        sessions.push(...rows.map(projectSession))
      }
      if (sessions.some(session => session.directory !== directory)) throw new Error("Engine session is outside the requested matter.")
      return projected(response, sessions)
    }
    if (url.pathname === "/api/session/active") {
      if (!record(value) || Object.values(value).some(item => !record(item) || item.type !== "running")) throw new Error("Engine returned invalid execution state.")
      const ids = Object.keys(value)
      if (!directory || ids.length > 32) return projected(response, Object.fromEntries(ids.map(id => [id, { type: "busy" }])))
      const scoped = await Promise.all(ids.map(async id => {
        const result = await transport(new URL(`/api/session/${encodeURIComponent(id)}`, url), { headers, signal: options.signal })
        if (!result.ok) throw new Error("Active engine session location is unavailable.")
        const session = projectSession(data(await result.json()))
        if (session.id !== id) throw new Error("Active engine session identity is invalid.")
        return session.directory === directory ? id : undefined
      }))
      return projected(response, Object.fromEntries(scoped.filter((id): id is string => id !== undefined).map(id => [id, { type: "busy" }])))
    }
    if (url.pathname === "/api/provider") {
      const modelsResponse = await transport(new URL("/api/model", url), { headers, signal: options.signal })
      if (!modelsResponse.ok) return modelsResponse
      const models = data(await modelsResponse.json())
      if (!Array.isArray(value) || !Array.isArray(models)) throw new Error("Engine returned invalid model catalog.")
      // Both endpoints expose the native available snapshot, including provider
      // connection filtering. Use selection IDs, not provider wire model IDs.
      const providers = value.filter(record).filter(provider => provider.activation !== "disabled")
      return projected(response, { all: providers.map(provider => ({ ...provider, models: Object.fromEntries(models.filter(record).filter(model => model.providerID === provider.id && model.enabled === true).map(model => [String(model.id), model])) })), connected: providers.map(provider => provider.id) })
    }
    return projected(response, value)
  }
  const wrapped = Object.assign(request, { preconnect: transport.preconnect }) as typeof fetch
  adapters.add(wrapped)
  return wrapped
}

function nativeModel(model: Record<string, unknown>) {
  return { providerID: model.providerID, id: model.id ?? model.modelID, ...(model.variant && model.variant !== "default" ? { variant: model.variant } : {}) }
}

function data(envelope: unknown): unknown {
  if (!record(envelope) || !("data" in envelope)) throw new Error("Engine returned an invalid v2 response envelope.")
  return envelope.data
}

function projected(response: Response, value: unknown) {
  return Response.json(value, { status: response.status })
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
