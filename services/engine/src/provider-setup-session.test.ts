import { expect, test } from "bun:test"
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Integration } from "@opencode/core/integration"
import { Provider } from "@opencode/core/provider"
import { Context, Effect, Logger } from "effect"
import { Database } from "bun:sqlite"
import { createProviderSetupSession, setupMemoryCredentials } from "./provider-setup-session"

test("setup credential selection and updates remain in the isolated store", async () => {
  const memory = setupMemoryCredentials()
  const id = Integration.ID.make("synthetic")
  await Effect.runPromise(Effect.gen(function* () {
    const first = yield* memory.service.create({ integrationID: id, value: { type: "key", key: "synthetic-first" } })
    const second = yield* memory.service.create({ integrationID: id, value: { type: "key", key: "synthetic-second" } })
    expect(memory.snapshot(id)?.value).toEqual(second.value)
    yield* memory.service.activate(first.id)
    expect(memory.snapshot(id)?.value).toEqual(first.value)
    yield* memory.service.update(first.id, { value: { type: "key", key: "synthetic-refreshed" } })
    const snapshot = memory.snapshot(id)!
    expect(snapshot.value).toEqual({ type: "key", key: "synthetic-refreshed" })
    if (snapshot.value.type === "key") (snapshot.value as { key: string }).key = "changed-copy"
    expect(memory.snapshot(id)?.value).toEqual({ type: "key", key: "synthetic-refreshed" })
    yield* memory.service.remove(first.id)
    expect(memory.snapshot(id)?.value).toEqual(second.value)
  }))
})

test("native setup graph stages credentials, applies OAuth provider semantics, and cleans its files", async () => {
  const legal = mkdtempSync(path.join(tmpdir(), "dochaus-setup-fixture-"))
  let directory = ""
  try {
    writeFileSync(path.join(legal, "opencode.json"), "{}")
    await Effect.runPromise(Effect.scoped(Effect.gen(function* () {
      const session = yield* createProviderSetupSession({
        legal, settings: {}, credential: {
          integrationID: Integration.ID.make("openai"), label: "Synthetic ChatGPT",
          value: { type: "oauth", methodID: Integration.MethodID.make("chatgpt-browser"), refresh: "synthetic-refresh", access: "synthetic-access", expires: Date.now() + 3_600_000 },
        },
      })
      directory = session.directory
      const provider = yield* session.provide(Provider.Service.use(service => service.get(Provider.ID.openai)))
      expect(provider?.settings?.baseURL).toBe("https://chatgpt.com/backend-api/codex")
      expect(session.credential("openai")?.label).toBe("Synthetic ChatGPT")
      // Exercise the real native key entry path; it writes only to memory.
      yield* session.provide(Integration.Service.use(service => service.connection.key({ integrationID: Integration.ID.make("openai"), key: "synthetic-api-key" })))
      expect(session.credential("openai")?.value).toEqual({ type: "key", key: "synthetic-api-key" })
      yield* Effect.sleep("10 millis")
      const replacement = yield* session.provide(Provider.Service.use(service => service.get(Provider.ID.openai)))
      expect(replacement?.settings?.baseURL).not.toBe("https://chatgpt.com/backend-api/codex")
      const integration = Context.get(session.context, Integration.Service)
      const integrationID = Integration.ID.make("synthetic-oauth")
      const methodID = Integration.MethodID.make("synthetic-code")
      let released = 0
      yield* integration.transform(editor => editor.method.update({
        integrationID, method: { type: "oauth", id: methodID, label: "Synthetic code" },
        authorize: () => Effect.gen(function* () {
          yield* Effect.addFinalizer(() => Effect.sync(() => { released++ }))
          return { mode: "code" as const, url: "https://example.test/synthetic", instructions: "Synthetic only",
            callback: (code: string) => Effect.succeed({ type: "oauth" as const, methodID, access: code, refresh: "synthetic-refresh", expires: Date.now() + 3_600_000 }),
          }
        }),
      }))
      const cancelled = yield* integration.oauth.connect({ integrationID, methodID })
      expect((yield* integration.oauth.status({ integrationID, attemptID: cancelled.attemptID })).status).toBe("pending")
      yield* integration.oauth.cancel({ integrationID, attemptID: cancelled.attemptID })
      expect(released).toBe(1)
      expect(session.credential(integrationID)).toBeUndefined()
      const completed = yield* integration.oauth.connect({ integrationID, methodID })
      yield* integration.oauth.complete({ integrationID, attemptID: completed.attemptID, code: "synthetic-code" })
      expect((yield* integration.oauth.status({ integrationID, attemptID: completed.attemptID })).status).toBe("complete")
      expect(session.credential(integrationID)?.value.type).toBe("oauth")
      const commandID = Integration.MethodID.make("synthetic-command")
      yield* integration.transform(editor => editor.method.update({
        integrationID, method: { type: "command", id: commandID, label: "Synthetic local command",
          command: [process.execPath, "-e", "process.stdout.write('synthetic-command-key')"],
        },
      }))
      const command = yield* integration.command.connect({ integrationID, methodID: commandID })
      let status = yield* integration.command.status({ integrationID, attemptID: command.attemptID })
      for (let polls = 0; status.status === "pending" && polls < 100; polls++) {
        yield* Effect.sleep("10 millis")
        status = yield* integration.command.status({ integrationID, attemptID: command.attemptID })
      }
      expect(status.status).toBe("complete")
      expect(session.credential(integrationID)?.value).toEqual({ type: "key", key: "synthetic-command-key" })
      expect(existsSync(path.join(directory, "setup.db"))).toBe(true)
      const db = new Database(path.join(directory, "setup.db"), { readonly: true })
      try { expect(db.query("SELECT count(*) AS count FROM credential").get()).toEqual({ count: 0 }) }
      finally { db.close() }
    })).pipe(Effect.provide(Logger.layer([]))))
    expect(directory).not.toBe("")
    expect(existsSync(directory)).toBe(false)
  } finally { rmSync(legal, { recursive: true, force: true }) }
}, 60_000)
