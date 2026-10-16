import { expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { OpenCode } from "@opencode/client"
import { startTrace } from "./support/native-trace"

test("native SSE records execution and counts each durable token event once", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-trace-"))
  const file = path.join(directory, "trace.jsonl")
  const event = {
    id: "evt_test", created: 1, type: "session.step.ended",
    durable: { aggregateID: "ses_test", seq: 1, version: 1 },
    data: { sessionID: "ses_test", assistantMessageID: "msg_test", finish: "stop", cost: 0,
      tokens: { input: 7, output: 3, reasoning: 2, cache: { read: 1, write: 0 } } },
  }
  const server = Bun.serve({ port: 0, fetch(request) {
    expect(new URL(request.url).pathname).toBe("/api/event")
    return new Response([event, event].map(item => `data: ${JSON.stringify(item)}\n\n`).join(""), {
      headers: { "Content-Type": "text/event-stream" },
    })
  } })
  const trace = startTrace(OpenCode.make({ baseUrl: server.url.href }), file)
  try {
    for (let attempt = 0; attempt < 100; attempt++) {
      if (await Bun.file(file).exists() && (await Bun.file(file).text()).includes("session.step.ended")) break
      await Bun.sleep(10)
    }
    const summary = await trace.stop()
    expect(summary.grand).toEqual({ input: 7, output: 3, reasoning: 2 })
    expect(summary.sessions).toBe(1)
    const records = (await Bun.file(file).text()).trim().split("\n").map(line => JSON.parse(line))
    expect(records.filter(record => record.type === "session.step.ended")).toHaveLength(2)
    expect(records.some(record => record.type === "trace.error")).toBe(false)
  } finally {
    server.stop(true)
    rmSync(directory, { recursive: true, force: true })
  }
})
