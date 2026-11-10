import { expect, test } from "bun:test"
import type { SessionCreateInput } from "@opencode/client"
import { createSession } from "../src/api/opencode/sessions"
import { routeAgent } from "../src/api/opencode/routing"
import { clientDirectories, type Client } from "../src/api/opencode/client"
import { isUserConversation } from "../../../dochaus/lib/session-visibility"
import { isConversation } from "../src/components/dashboard/data"

function fixture() {
  const created: SessionCreateInput[] = [], removed: string[] = []
  const client = {
    session: {
      create: async (input: SessionCreateInput) => { created.push(input); return { id:"automatic", ...input } },
      switchAgent: async () => {}, switchModel: async () => {}, prompt: async () => {}, wait: async () => {}, interrupt: async () => {},
      remove: async ({sessionID}:{sessionID:string}) => {removed.push(sessionID)},
    },
    message:{list:async()=>({data:[{type:"assistant",id:"answer",agent:"router",time:{created:1,completed:2},content:[{type:"text",text:'{"route":"qa"}'}]}],cursor:{}})},
  } as unknown as Client
  clientDirectories.set(client,"/synthetic/matter")
  return {client,created,removed}
}

test("actual router producer creates an explicitly owned session excluded by conversation and dashboard predicates", async () => {
  const {client,created,removed}=fixture()
  expect(await routeAgent({client,smallModel:"fixture/fast",candidates:[{name:"qa",description:"Answer"}],history:[],text:"Explain this clause"})).toEqual({name:"qa"})
  expect(created).toHaveLength(1)
  expect(created[0]).toMatchObject({location:{directory:"/synthetic/matter"},metadata:{dochaus_role:"chat-router"}})
  expect(created[0]).not.toHaveProperty("parentID")
  const session={metadata:created[0]!.metadata}
  expect(isUserConversation(session)).toBe(false)
  expect(isConversation(session)).toBe(false)
  expect(removed).toEqual(["automatic"])
})

test("ordinary session producer preserves human titles and metadata-free legacy visibility", async () => {
  const {client,created}=fixture(), title="doc.haus grid: My own discussion"
  await createSession(client,title,1000,[{action:"*",resource:"*",effect:"deny"}])
  expect(created[0]).toEqual({title,location:{directory:"/synthetic/matter"},permissions:[{action:"*",resource:"*",effect:"deny"}]})
  const session={metadata:created[0]!.metadata}
  expect(isUserConversation(session)).toBe(true)
  expect(isConversation(session)).toBe(true)
})
