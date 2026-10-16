import { afterEach, expect, test } from "bun:test"
import type { SessionDomain } from "@opencode/plugin/promise/session"
import { mkdtemp, realpath, rm, symlink } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { nativeSessionReader } from "./native-session-reader"
import { engineV2Fetch } from "./engine-v2"
import { isProjectedUserAdmission } from "./engine-v2-projection"

const directories: string[] = []
afterEach(async () => { await Promise.all(directories.splice(0).map(directory => rm(directory, { recursive: true, force: true }))) })
async function fixture() {
  const directory = await realpath(await mkdtemp(path.join(tmpdir(), "native-session-reader-")))
  directories.push(directory)
  const session: Awaited<ReturnType<SessionDomain["get"]>> = { id: "ses_native_reader", projectID: "global", location: { directory }, agent: "draft", model: { providerID: "neutral-provider", id: "neutral-model" }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 0, updated: 5 }, fork: { sessionID: "ses_native_parent", boundary: { type: "through", messageID: "msg_parent_boundary" } } }
  const messages: Awaited<ReturnType<SessionDomain["messages"]>> = [
    { id: "msg_exact_request", type: "user", text: "Original exact request", metadata: { dochaus_tools: { write: false } }, time: { created: 0 } },
    ...Array.from({ length: 201 }, (_, index) => ({ id: `msg_retained_${index}`, type: "synthetic" as const, text: `Retained neutral instruction ${index}`, time: { created: index + 1 } })),
    { id: "msg_compaction", type: "compaction", status: "completed", summary: "Neutral summary", recent: "Recent material", reason: "manual", time: { created: 203, completed: 204 } },
    { id: "msg_native_parts", type: "assistant", agent: "draft", model: session.model!, time: { created: 205, completed: 206 }, finish: "stop", content: [{ type: "text", text: "First exact part" }, { type: "text", text: "Second exact part" }] },
    { id: "msg_idle", type: "idle", outcome: "succeeded", time: { created: 207 } },
  ]
  return { directory, session, messages }
}

test("native full-history projection equals paginated wire projection without HTTP or context reads", async () => {
  const f = await fixture()
  let metadata = 0, histories = 0
  const reader = await nativeSessionReader(f.directory, {
    get: async input => { expect(input).toEqual({ sessionID: f.session.id }); metadata++; return f.session },
    messages: async input => { expect(input).toEqual({ sessionID: f.session.id }); histories++; return f.messages },
  })
  const read = await reader.read(f.session.id)
  expect({ metadata, histories }).toEqual({ metadata: 1, histories: 1 })
  expect(read.session.fork).toEqual(f.session.fork)
  expect(read.messages).toHaveLength(204)
  expect(read.messages.slice(0, 202).map(message => message.info.id)).toEqual(["msg_exact_request", ...Array.from({ length: 201 }, (_, index) => `msg_retained_${index}`)])
  expect(isProjectedUserAdmission(read.messages[0])).toBe(true)
  expect(isProjectedUserAdmission(read.messages[1])).toBe(false)
  expect(read.messages[0].info.tools).toEqual({ write: false })
  expect(read.messages[202].info).toMatchObject({ id: "msg_compaction", summary: true })
  expect(read.messages[203].parts.map(part => [part.id, part.type, part.text])).toEqual([["msg_native_parts_0", "text", "First exact part"], ["msg_native_parts_1", "text", "Second exact part"]])
  const gets: string[] = []
  const transport = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(input instanceof Request ? input.url : String(input))
    gets.push(url.pathname)
    if (!url.pathname.endsWith("/message")) return Response.json({ data: f.session })
    const reversed = f.messages.toReversed()
    const offset = Number(url.searchParams.get("cursor") ?? "0")
    const rows = reversed.slice(offset, offset + 200)
    return Response.json({ data: rows, cursor: rows.length ? { next: String(offset + rows.length) } : {} })
  }, { preconnect: fetch.preconnect })
  const response = await engineV2Fetch(transport)(`http://127.0.0.1/api/session/${f.session.id}/message`, { headers: { "x-opencode-directory": f.directory } })
  expect(await response.json()).toEqual(read.messages)
  expect(gets).toEqual([`/api/session/${f.session.id}`, ...Array.from({ length: 3 }, () => `/api/session/${f.session.id}/message`)])
})

test("metadata-only native get never reads history and accepts only canonical matter identity", async () => {
  const f = await fixture()
  const alias = f.directory + "-alias"
  directories.push(alias)
  await symlink(f.directory, alias)
  const reader = await nativeSessionReader(alias, { get: async () => f.session, messages: async () => { throw new Error("Metadata must not fetch history") } })
  expect((await reader.get(f.session.id)).id).toBe(f.session.id)
})

test("foreign, missing or mismatched native session identity denies before history", async () => {
  const f = await fixture(), foreign = await fixture()
  let histories = 0
  for (const value of [{ ...f.session, location: { directory: foreign.directory } }, { ...f.session, id: "ses_foreign_identity" }, { ...f.session, location: { directory: path.join(f.directory, "missing") } }]) {
    const reader = await nativeSessionReader(f.directory, { get: async () => value, messages: async () => { histories++; return f.messages } })
    await expect(reader.read(f.session.id)).rejects.toThrow()
  }
  const missing = await nativeSessionReader(f.directory, { get: async () => { throw new Error("Native session not found") }, messages: async () => { histories++; return f.messages } })
  await expect(missing.read(f.session.id)).rejects.toThrow("Native session not found")
  expect(histories).toBe(0)
})

test("native history failure remains failed and existing latest native-row policy precedes projection", async () => {
  const f = await fixture()
  const failed = await nativeSessionReader(f.directory, { get: async () => f.session, messages: async () => { throw new Error("Native durable history decode failed") } })
  await expect(failed.read(f.session.id)).rejects.toThrow("Native durable history decode failed")
  const reader = await nativeSessionReader(f.directory, { get: async () => f.session, messages: async () => f.messages })
  const limited = await reader.read(f.session.id, 2)
  expect(limited.messages.map(message => message.info.id)).toEqual(["msg_native_parts"])
  expect((await reader.read(f.session.id)).messages.map(message => message.info.id)).toContain("msg_exact_request")
})
