import { expect, test } from "bun:test"
import type { WordChatMaterial } from "../../../dochaus/lib/word-chat-contract"
import type { WordChatApi } from "./chat-client"
import { chatAttachmentRequestHandler } from "./chat-attachments"
import { CHAT_UPLOAD_MAX_BASE64, DESK_MESSAGE_LIMIT, deskMessage, isDeskRequest, isDeskView, readDeskMessage, type DeskRequest } from "./desk-protocol"

const source: WordChatMaterial = { kind: "source", id: "one", hash: "a".repeat(64), name: "Instructions.docx", provenance: "saved-version" }
const uploaded: WordChatMaterial = { kind: "upload", id: "new-upload", hash: "b".repeat(64), name: "Notes.txt", provenance: "request-upload" }
const ref = ({ kind, id, hash }: WordChatMaterial) => ({ kind, id, hash })

test("only validated attachment uploads exceed the existing one MiB bridge boundary", () => {
  const data = "A".repeat(Math.ceil(DESK_MESSAGE_LIMIT / 3) * 4)
  const request = { action: "chat-upload" as const, name: "Notes.txt", data, mode: "request" as const }
  const encoded = deskMessage("channel", { kind: "request", id: "upload", request })
  expect(encoded.length).toBeGreaterThan(DESK_MESSAGE_LIMIT)
  expect(readDeskMessage(encoded)?.kind).toBe("request")
  expect(() => deskMessage("channel", { kind: "response", id: "read", ok: true, value: data })).toThrow("too large")
  const parsed = JSON.parse(encoded)
  expect(readDeskMessage(JSON.stringify({ ...parsed, kind: "response", request: undefined, ok: true, value: data }))).toBeUndefined()
  expect(readDeskMessage(JSON.stringify({ ...parsed, request: { ...request, conversationId: "another-conversation" } }))).toBeUndefined()
  expect(isDeskRequest({ ...request, data: "A".repeat(CHAT_UPLOAD_MAX_BASE64) })).toBe(false) // decoded bytes exceed 16 MiB without required padding
  expect(isDeskRequest({ ...request, data: "A".repeat(CHAT_UPLOAD_MAX_BASE64 - 2) + "==" })).toBe(true)
  for (const changed of [{ name: "../Notes.txt" }, { name: "Notes.exe" }, { data: "AA?=" }, { mode: "elsewhere" }])
    expect(isDeskRequest({ ...request, ...changed })).toBe(false)
  expect(isDeskView("attachments")).toBe(false)
})

test("picker requests stay bound to the parent conversation and validate a fresh material identity", async () => {
  const calls: { action: string; input?: object }[] = [], attached: WordChatMaterial[] = []
  let catalog = [source]
  const api: WordChatApi = {
    async request<T>(action: Parameters<WordChatApi["request"]>[0], input?: object) { calls.push({ action, input }); return { materials: catalog, matterName: "Synthetic matter", documentName: "Draft.docx", token: "must-not-cross" } as T },
    async upload() { throw new Error("Unexpected upload") },
  }
  const handler = chatAttachmentRequestHandler(api, { conversationId: "parent-conversation", attached: [], onAttach: material => { attached.push(material) } })
  const result = await handler.request({ action: "chat-materials" })
  expect(result).toEqual({ materials: [source], matterName: "Synthetic matter", documentName: "Draft.docx", attached: [] })
  await handler.request({ action: "chat-attach", material: ref(source) })
  await handler.request({ action: "chat-attach", material: ref(source) })
  expect(attached).toEqual([source])
  expect(calls.every(call => JSON.stringify(call.input) === JSON.stringify({ conversationId: "parent-conversation" }))).toBe(true)
  const before = calls.length
  for (const request of [{ action: "selection" }, { action: "chat-materials", conversationId: "foreign" }, { action: "chat-attach", material: { ...ref(source), name: "Forged name" } }])
    await expect(handler.request(request as DeskRequest)).rejects.toThrow("Choose a file")
  expect(calls).toHaveLength(before)
  catalog = [{ ...source, hash: "c".repeat(64) }]
  await expect(handler.request({ action: "chat-attach", material: ref(source) })).rejects.toThrow("changed")
  expect(attached).toEqual([source])
  handler.dispose()
  await expect(handler.request({ action: "chat-materials" })).rejects.toThrow("closed")
})

test("binary uploads use the captured conversation and selected destination, then attach exactly once", async () => {
  const attached: WordChatMaterial[] = [], uploads: { conversation: string; mode: string; file: File }[] = []
  const api: WordChatApi = {
    async request() { throw new Error("Unexpected read") },
    async upload<T>(conversation: string, file: File, mode: "request" | "matter") { uploads.push({ conversation, file, mode }); return { material: uploaded } as T },
  }
  const handler = chatAttachmentRequestHandler(api, { conversationId: "parent-conversation", attached: [], onAttach: material => { attached.push(material) } })
  const data = btoa(String.fromCharCode(0, 255, 65, 10))
  expect(await handler.request({ action: "chat-upload", name: "Notes.txt", data, mode: "matter" })).toEqual({ material: uploaded })
  expect(uploads).toHaveLength(1)
  expect(uploads[0].conversation).toBe("parent-conversation")
  expect(uploads[0].mode).toBe("matter")
  expect(uploads[0].file.name).toBe("Notes.txt")
  expect([...new Uint8Array(await uploads[0].file.arrayBuffer())]).toEqual([0, 255, 65, 10])
  expect(attached).toEqual([uploaded])
})

test("closing a dialog fences a late upload response and parallel changes cannot duplicate an attachment", async () => {
  let resolve!: (value: { material: WordChatMaterial }) => void
  const pending = new Promise<{ material: WordChatMaterial }>(done => { resolve = done })
  let callbacks = 0, uploads = 0
  const api: WordChatApi = {
    async request() { throw new Error("Unexpected read") },
    async upload<T>() { uploads++; return await pending as T },
  }
  const handler = chatAttachmentRequestHandler(api, { conversationId: "parent", attached: [], onAttach() { callbacks++ } })
  const input = { action: "chat-upload" as const, name: "Notes.txt", data: "YWJj", mode: "request" as const }
  const request = handler.request(input)
  await expect(handler.request(input)).rejects.toThrow("Wait")
  expect(uploads).toBe(1)
  handler.dispose()
  resolve({ material: uploaded })
  await expect(request).rejects.toThrow("closed")
  expect(callbacks).toBe(0)
})
