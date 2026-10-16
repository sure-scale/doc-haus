import { expect, test } from "bun:test"
import { OpenCode } from "@opencode/client"

test("native session forms carry directory, answer selections and dismissal through their own APIs", async () => {
  const requests: { method: string; path: string; directory: string | null; body?: unknown }[] = []
  const form = {
    id: "form-intake",
    sessionID: "session-library",
    title: "Questions",
    metadata: { kind: "question" },
    fields: [{ key: "q0", type: "string", custom: true, options: [{ value: "Neutral", label: "Neutral" }] }],
  }
  const server = Bun.serve({
    port: 0,
    fetch: async (request) => {
      requests.push({
        method: request.method,
        path: new URL(request.url).pathname,
        directory: request.headers.get("x-opencode-directory"),
        body: request.method === "POST" ? await request.json() : undefined,
      })
      return request.method === "GET" ? Response.json({ data: [form] }) : new Response(null, { status: 204 })
    },
  })
  try {
    const client = OpenCode.make({ baseUrl: server.url.href, headers: { "x-opencode-directory": "/library" } })
    expect(await client.session.form.list({ sessionID: form.sessionID })).toEqual([form])
    await client.session.form.reply({
      sessionID: form.sessionID,
      formID: form.id,
      answer: { q0: "Neutral", q1: ["Keep", "Other clause"] },
    })
    await client.session.form.cancel({ sessionID: form.sessionID, formID: form.id })
    expect(requests).toEqual([
      { method: "GET", path: "/api/session/session-library/form", directory: "/library", body: undefined },
      {
        method: "POST",
        path: "/api/session/session-library/form/form-intake/reply",
        directory: "/library",
        body: { answer: { q0: "Neutral", q1: ["Keep", "Other clause"] } },
      },
      {
        method: "DELETE",
        path: "/api/session/session-library/form/form-intake",
        directory: "/library",
        body: undefined,
      },
    ])
  } finally {
    server.stop(true)
  }
})
