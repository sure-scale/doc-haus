import { expect, test } from "bun:test"
import { generateConversationTitle } from "./chat-session-title"

test("conversation naming delegates to the native title action and can name a scoped container", async () => {
  const calls: Array<[string, string, string | undefined]> = []
  const titles = new Map([["child", "Request placeholder"], ["root", "Conversation"]])
  const title = await generateConversationTitle({ sessionID: "child", targetSessionID: "root", transport: {
    get: async id => ({ title: titles.get(id) }),
    rename: async (id, value, expected) => { calls.push([id, value, expected]); if (expected === undefined || titles.get(id) === expected) titles.set(id, value || "Delivery clause review") },
  } })
  expect(calls).toEqual([["child", "", undefined], ["root", "Delivery clause review", "Conversation"]])
  expect(title).toBe("Delivery clause review")
  expect(titles.get("root")).toBe(title)
})

test("a user's concurrent conversation rename wins over a generated child title", async () => {
  const calls: Array<[string, string, string | undefined]> = []
  const titles = new Map([["child", "Request placeholder"], ["root", "Conversation"]])
  await generateConversationTitle({ sessionID: "child", targetSessionID: "root", transport: {
    get: async id => ({ title: titles.get(id) }),
    rename: async (id, value, expected) => {
      calls.push([id, value, expected])
      if (id === "child") { titles.set(id, "Delivery clause review"); titles.set("root", "Lawyer's own title") }
      else if (titles.get(id) === expected) titles.set(id, value)
    },
  } })
  expect(calls).toEqual([["child", "", undefined], ["root", "Delivery clause review", "Conversation"]])
  expect(titles.get("root")).toBe("Lawyer's own title")
})

test("a no-op native title action does not copy the child placeholder", async () => {
  const calls: Array<[string, string]> = []
  const titles = new Map([["child", "Word response"], ["root", "Word conversation"]])
  const title = await generateConversationTitle({ sessionID: "child", targetSessionID: "root", transport: {
    get: async id => ({ title: titles.get(id) }),
    rename: async (id, value) => { calls.push([id, value]) },
  } })
  expect(title).toBeUndefined()
  expect(calls).toEqual([["child", ""]])
  expect(titles.get("root")).toBe("Word conversation")
})
