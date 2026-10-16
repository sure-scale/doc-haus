import { describe, expect } from "bun:test"
import { deepStrictEqual } from "node:assert"
import { Agent } from "@opencode/schema/agent"
import { Model } from "@opencode/schema/model"
import { Provider } from "@opencode/schema/provider"
import { Money } from "@opencode/schema/money"
import { OpenAIChat } from "@opencode/ai/protocols"
import { fromPromise } from "@opencode/plugin/promise/adapter"
import { Bus } from "@opencode/core/bus"
import { Database } from "@opencode/core/database/database"
import { AppNodeBuilder } from "@opencode/core/effect/app-node-builder"
import { LayerNode } from "@opencode/util/effect/layer-node"
import { Location } from "@opencode/core/location"
import { PluginHost } from "@opencode/core/plugin/host"
import { Session } from "@opencode/core/session"
import { SessionEvent } from "@opencode/core/session/event"
import { SessionExecution } from "@opencode/core/session/execution"
import { SessionInbox } from "@opencode/core/session/inbox"
import { SessionMessage } from "@opencode/core/session/message"
import { SessionModelRequest } from "@opencode/core/session/model-request"
import { SessionModelTransport } from "@opencode/core/session/model-transport"
import { SessionRunnerModel } from "@opencode/core/session/runner/model"
import type { SessionStore } from "@opencode/core/session/store"
import { Effect, Exit, Schema } from "effect"
import { sql } from "drizzle-orm"
import { testEffect } from "./lib/effect"
import { offlineModels } from "./fixture/models"
import { PluginTestGraph, PluginTestReplacements } from "./plugin/fixture"

const it = testEffect(
  AppNodeBuilder.build(LayerNode.group([PluginTestGraph, Database.node]), [
    ...PluginTestReplacements,
    Bus.node.replace(Bus.configured({ persist: true })),
    SessionExecution.node.replace(SessionExecution.noopLayer),
    offlineModels,
  ]),
)
const model = Model.Ref.make({ id: Model.ID.make("fixture-model"), providerID: Provider.ID.make("fixture-provider") })
const usage = { cost: Money.USD.zero, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } } }

describe("Plugin native full session history", () => {
  it.effect("exposes complete oldest-first history through real host and Promise codecs across compaction and fork", () =>
    Effect.gen(function* () {
      const sessions = yield* Session.Service
      const location = yield* Location.Service
      const bus = yield* Bus.Service
      const { db } = yield* Database.Service
      const parent = yield* sessions.create({ location: Location.Ref.make({ directory: location.directory }), agent: Agent.ID.make("build"), model })
      const promptID = SessionMessage.ID.make("msg_history_original_admission")
      yield* sessions.prompt({ sessionID: parent.id, id: promptID, text: "Original exact request", resume: false })
      expect((yield* sessions.messages({ sessionID: parent.id })).map(message => message.id)).not.toContain(promptID)
      yield* SessionInbox.promote(db, bus, parent.id, "steer")
      const notes = Array.from({ length: 201 }, (_, index) => ({ id: SessionMessage.ID.make(`msg_history_note_${index}`), text: `Neutral retained note ${index}` }))
      for (const note of notes) {
        yield* sessions.synthetic({ sessionID: parent.id, ...note, resume: false })
        yield* SessionInbox.promote(db, bus, parent.id, "steer")
      }
      yield* bus.publish(SessionEvent.Compaction.Started, { sessionID: parent.id, reason: "manual", recent: "Retained recent material" })
      yield* bus.publish(SessionEvent.Compaction.Ended, { sessionID: parent.id, reason: "manual", text: "Neutral compaction summary", recent: "Retained recent material" })
      const laterID = SessionMessage.ID.make("msg_history_later_admission")
      yield* sessions.prompt({ sessionID: parent.id, id: laterID, text: "Later exact request", resume: false })
      yield* SessionInbox.promote(db, bus, parent.id, "steer")
      const assistantID = SessionMessage.ID.make("msg_history_assistant")
      yield* bus.publish(SessionEvent.Step.Started, { sessionID: parent.id, assistantMessageID: assistantID, agent: Agent.ID.make("build"), model, started: 0 })
      for (const [ordinal, text] of ["First exact part", "Second exact part"].entries()) {
        yield* bus.publish(SessionEvent.Text.Started, { sessionID: parent.id, assistantMessageID: assistantID, ordinal })
        yield* bus.publish(SessionEvent.Text.Ended, { sessionID: parent.id, assistantMessageID: assistantID, ordinal, text })
      }
      yield* bus.publish(SessionEvent.Step.Ended, { sessionID: parent.id, assistantMessageID: assistantID, finish: "stop", ...usage })
      const forked = yield* sessions.fork({ sessionID: parent.id })
      const reads: SessionStore.MessagesInput[] = []
      const native = Session.Service.of({ ...sessions, messages: input => { reads.push(input); return sessions.messages(input) } })
      const host = yield* PluginHost.make({ list: () => Effect.succeed([]) }).pipe(Effect.provideService(Session.Service, native))
      expect((yield* host.session.get({ sessionID: parent.id })).id).toBe(parent.id)
      expect(reads).toEqual([])
      const rows = yield* host.session.messages({ sessionID: parent.id })
      expect(reads).toEqual([{ sessionID: parent.id, order: "asc" }])
      expect(rows).toHaveLength(205)
      expect(rows.slice(0, 202).map(message => message.id)).toEqual([promptID, ...notes.map(note => note.id)])
      expect(rows.filter(message => message.type === "user" || message.type === "synthetic").map(message => message.text)).toEqual(["Original exact request", ...notes.map(note => note.text), "Later exact request"])
      expect(rows[202]).toMatchObject({ type: "compaction", summary: "Neutral compaction summary", status: "completed" })
      expect(rows[203].id).toBe(laterID)
      expect(rows[204]).toMatchObject({ id: assistantID, type: "assistant", content: [{ type: "text", text: "First exact part" }, { type: "text", text: "Second exact part" }] })
      expect((yield* host.session.context({ sessionID: parent.id })).map(message => message.id)).not.toContain(promptID)
      const forkRows = yield* host.session.messages({ sessionID: forked.id })
      expect(forkRows).toHaveLength(205)
      expect(forkRows.map(message => message.type)).toEqual(rows.map(message => message.type))
      expect(forkRows.filter(message => message.type === "user" || message.type === "synthetic").map(message => message.text)).toEqual(["Original exact request", ...notes.map(note => note.text), "Later exact request"])
      expect(forkRows.map(message => message.id)).not.toEqual(rows.map(message => message.id))
      expect(forked.fork?.sessionID).toBe(parent.id)
      expect(new Set(forkRows.map(message => message.id)).size).toBe(205)
      // These user/synthetic/compaction/assistant rows have identical native and public wire fields.
      // The real Promise adapter still applies its public codec; no HTTP-only location row is seeded.
      const encode = Schema.encodeSync(Schema.Array(SessionMessage.Info))
      yield* fromPromise({
        id: "full-history-codec-proof",
        async setup(context) {
          expect((await context.session.get({ sessionID: parent.id })).id).toBe(parent.id)
          expect(reads).toHaveLength(2)
          deepStrictEqual(await context.session.messages({ sessionID: parent.id }), encode(rows))
          deepStrictEqual(await context.session.messages({ sessionID: forked.id }), encode(forkRows))
          expect(reads).toEqual([parent.id, forked.id, parent.id, forked.id].map(sessionID => ({ sessionID, order: "asc" })))
        },
      }).effect(host)
    }),
  )

  it.effect("propagates a native full-history decode failure through Promise context preparation before the stream boundary", () =>
    Effect.gen(function* () {
      const sessions = yield* Session.Service
      const location = yield* Location.Service
      const bus = yield* Bus.Service
      const { db } = yield* Database.Service
      const session = yield* sessions.create({ location: Location.Ref.make({ directory: location.directory }), agent: Agent.ID.make("build"), model })
      const id = SessionMessage.ID.make("msg_history_corrupt")
      yield* sessions.prompt({ sessionID: session.id, id, text: "Neutral admitted request", resume: false })
      yield* SessionInbox.promote(db, bus, session.id, "steer")
      yield* db.run(sql`update session_message set data = '{"time":{"created":0}}' where id = ${id}`).pipe(Effect.orDie)
      const host = yield* PluginHost.make({ list: () => Effect.succeed([]) })
      const expected = { _tag: "Session.MessageDecodeError", sessionID: session.id, messageID: id }
      expect(yield* host.session.messages({ sessionID: session.id }).pipe(Effect.flip)).toMatchObject(expected)
      yield* fromPromise({
        id: "failed-native-read",
        async setup(context) {
          await context.session.hook("context", async event => { await context.session.messages({ sessionID: event.sessionID }) })
        },
      }).effect(host)
      const requests = yield* SessionModelRequest.Service.pipe(Effect.provide(SessionModelRequest.layer))
      let streamBoundary = 0
      const prepared = yield* requests.primary({ session, agent: Agent.ID.make("build"), model: SessionRunnerModel.resolved(OpenAIChat.route.model({ id: "fixture-model", provider: "fixture-provider" }), { capabilities: { tools: true, input: ["text"], output: ["text"] }, cost: [], limit: { context: 4096, output: 1024 } }), system: [], messages: [] }).pipe(Effect.andThen(Effect.sync(() => { streamBoundary++ })), Effect.exit)
      expect(Exit.isFailure(prepared)).toBe(true)
      expect(streamBoundary).toBe(0)
    }).pipe(Effect.provideService(SessionModelTransport.Service, SessionModelTransport.Service.of({ bind: () => ({ execute: () => Effect.die("Unexpected model transport execution") }), close: () => Effect.void, closeAll: Effect.void }))),
  )
})
