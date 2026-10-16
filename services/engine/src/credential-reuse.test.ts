import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { Integration } from "@opencode/schema/integration"
import { mkdtempSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { createReusedCredential, credentialReuse, credentialReuseHandler } from "./credential-reuse"

const roots: string[] = []
afterEach(() => { for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true }) })
function fixture(filename = "opencode-local.db") {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-auth-reuse-"))
  roots.push(directory)
  const file = path.join(directory, filename)
  const db = new Database(file)
  db.exec(`CREATE TABLE credential (id TEXT PRIMARY KEY, integration_id TEXT, label TEXT, value TEXT${filename === "opencode.db" ? ", connector_id TEXT, method_id TEXT, active INTEGER, time_created INTEGER, time_updated INTEGER" : ""})`)
  const key = { type: "key", key: "synthetic-reuse-key", metadata: { region: "synthetic" }, configuration: { tenant: "test" } }
  const oauth = { type: "oauth", methodID: "device", refresh: "synthetic-refresh", access: "synthetic-access", expires: 4102444800000, metadata: { accountID: "synthetic-account", nested: { enabled: true } } }
  db.query("INSERT INTO credential (id, integration_id, label, value) VALUES (?, ?, ?, ?)").run("source-key", "opencode-go", "Test key", JSON.stringify(key))
  db.query("INSERT INTO credential (id, integration_id, label, value) VALUES (?, ?, ?, ?)").run("source-oauth", "opencode", "Test OAuth", JSON.stringify(oauth))
  db.close()
  const imported: unknown[] = []
  const create = async (input: unknown) => { imported.push(input); return { id: "new-native-credential" } }
  const service = credentialReuse({ directory, targetDirectory: path.join(directory, "target"), create })
  return { directory, file, key, oauth, imported, service, create }
}

test("discovery returns metadata only without source writes or sidecars", () => {
  const f = fixture()
  const before = readFileSync(f.file)
  const listing = f.service.list()
  expect(listing.sources[0].connections).toHaveLength(2)
  const wire = JSON.stringify(listing)
  for (const secret of [f.key.key, f.oauth.refresh, f.oauth.access, "source-key", "source-oauth", f.directory]) expect(wire).not.toContain(secret)
  expect(f.imported).toEqual([])
  expect(readFileSync(f.file)).toEqual(before)
  expect(readdirSync(f.directory)).toEqual(["opencode-local.db"])
})

test("current opencode.db is independently discoverable and copies only explicitly selected values", async () => {
  const f = fixture("opencode.db")
  const current = new Database(f.file)
  current.query("UPDATE credential SET integration_id = ?, value = ?, method_id = ? WHERE id = ?").run(
    "openai", JSON.stringify({ ...f.oauth, methodID: "chatgpt-browser" }), "synthetic-column-method-not-used", "source-oauth",
  )
  current.close()
  // A different native release may have created the local DB without the
  // credential table. That must not hide the current store or legacy sign-ins.
  const local = new Database(path.join(f.directory, "opencode-local.db"))
  local.exec("CREATE TABLE unrelated (id TEXT)")
  local.close()
  writeFileSync(path.join(f.directory, "auth.json"), JSON.stringify({ other: { type: "api", key: "synthetic-unselected-legacy" } }))
  const files = readdirSync(f.directory)
  const before = files.map(file => readFileSync(path.join(f.directory, file)))
  const listing = f.service.list()
  const source = listing.sources.find(source => source.label === "OpenCode connections")!
  expect(source.status).toBe("ready")
  expect(source.connections.map(connection => connection.integrationID).sort()).toEqual(["openai", "opencode-go"])
  expect(listing.sources.find(source => source.label === "OpenCode local connections")?.status).toBe("unavailable")
  expect(listing.sources.find(source => source.label === "OpenCode legacy sign-ins")?.status).toBe("ready")
  const wire = JSON.stringify(listing)
  for (const secret of [f.key.key, f.oauth.access, f.oauth.refresh, "synthetic-unselected-legacy", "source-key", "source-oauth", f.directory, "synthetic-column-method-not-used"]) expect(wire).not.toContain(secret)
  expect(f.imported).toEqual([])
  await f.service.importSelected({ sourceId: source.id, connectionId: source.connections.find(connection => connection.integrationID === "opencode-go")!.id })
  expect(f.imported).toEqual([{ integrationID: "opencode-go", label: "Test key", value: f.key }])
  await f.service.importSelected({ sourceId: source.id, connectionId: source.connections.find(connection => connection.integrationID === "openai")!.id })
  expect(f.imported[1]).toEqual({ integrationID: "openai", label: "Test OAuth", value: { ...f.oauth, methodID: "chatgpt-browser" } })
  expect(readdirSync(f.directory)).toEqual(files)
  files.forEach((file, index) => expect(readFileSync(path.join(f.directory, file))).toEqual(before[index]))
})

test("selection imports only one connection, preserving OAuth and API-key metadata", async () => {
  const f = fixture()
  const before = readFileSync(f.file)
  const source = f.service.list().sources[0]
  const oauth = source.connections.find(connection => connection.integrationID === "opencode")!
  const result = await f.service.importSelected({ sourceId: source.id, connectionId: oauth.id })
  expect(f.imported).toEqual([{ integrationID: "opencode", label: "Test OAuth", value: f.oauth }])
  expect(result.connection).toEqual({ id: "new-native-credential", integrationID: Integration.ID.make("opencode"), label: "Test OAuth", authType: "oauth" })
  expect(JSON.stringify(result)).not.toContain(f.oauth.access)
  const key = source.connections.find(connection => connection.integrationID === "opencode-go")!
  await f.service.importSelected({ sourceId: source.id, connectionId: key.id })
  expect(f.imported[1]).toEqual({ integrationID: "opencode-go", label: "Test key", value: f.key })
  expect(readFileSync(f.file)).toEqual(before)
})

test("legacy OAuth uses native migration semantics and never imports well-known discovery", async () => {
  const f = fixture()
  writeFileSync(path.join(f.directory, "auth.json"), JSON.stringify({
    openai: { type: "oauth", refresh: "synthetic-legacy-refresh", access: "synthetic-legacy-access", expires: 4102444800000, accountId: "synthetic-account", enterpriseUrl: "https://synthetic.invalid" },
    other: { type: "api", key: "synthetic-unselected" },
    "https://discovery.invalid": { type: "wellknown", key: "synthetic-command", token: "synthetic-token" },
  }))
  const before = readFileSync(path.join(f.directory, "auth.json"))
  const source = f.service.list().sources.find(source => source.label === "OpenCode legacy sign-ins")!
  expect(source.connections).toHaveLength(2)
  expect(source.unavailableCount).toBe(1)
  await f.service.importSelected({ sourceId: source.id, connectionId: source.connections.find(connection => connection.integrationID === "openai")!.id })
  expect(f.imported).toEqual([{ integrationID: "openai", label: "OAuth", value: {
    type: "oauth", methodID: "chatgpt-browser", refresh: "synthetic-legacy-refresh", access: "synthetic-legacy-access", expires: 4102444800000,
    metadata: { accountID: "synthetic-account", enterpriseUrl: "https://synthetic.invalid" },
  } }])
  expect(readFileSync(path.join(f.directory, "auth.json"))).toEqual(before)
})

test("unavailable source does not hide usable sign-ins from another source", () => {
  const f = fixture()
  writeFileSync(f.file, "synthetic-invalid-sqlite")
  writeFileSync(path.join(f.directory, "auth.json"), JSON.stringify({ "opencode-go": { type: "api", key: "synthetic-legacy-key" } }))
  const sources = f.service.list().sources
  expect(sources[0].status).toBe("unavailable")
  expect(sources[1].connections[0].integrationID).toBe("opencode-go")
  expect(JSON.stringify(sources)).not.toContain("synthetic-invalid-sqlite")
})

test("unknown, mismatched, stale and path-bearing selections never import", async () => {
  const f = fixture()
  const source = f.service.list().sources[0]
  for (const input of [undefined, {}, { sourceId: source.id, connectionId: "unknown" },
    { sourceId: "other-source", connectionId: source.connections[0].id },
    { sourceId: source.id, connectionId: source.connections[0].id, path: f.directory }]) await expect(f.service.importSelected(input)).rejects.toThrow()
  const db = new Database(f.file)
  db.exec("DELETE FROM credential")
  db.close()
  await expect(f.service.importSelected({ sourceId: source.id, connectionId: source.connections[0].id })).rejects.toThrow("Could not reuse")
  expect(f.imported).toEqual([])
})

test("parse and writer failures cannot expose values or source paths", async () => {
  const f = fixture()
  const source = f.service.list().sources[0]
  const db = new Database(f.file)
  db.query("UPDATE credential SET value = ?").run(JSON.stringify({ type: "key", key: 12, metadata: { secret: "synthetic-error-secret" } }))
  db.close()
  await expect(f.service.importSelected({ sourceId: source.id, connectionId: source.connections[0].id })).rejects.toThrow("Could not reuse the selected sign-in.")
  expect(f.imported).toEqual([])
  const failing = credentialReuse({ directory: f.directory, targetDirectory: path.join(f.directory, "target"), create: async () => { throw new Error("synthetic-native-error-secret") } })
  const restore = new Database(f.file)
  restore.query("UPDATE credential SET value = ?").run(JSON.stringify(f.key))
  restore.close()
  const source2 = failing.list().sources[0]
  await expect(failing.importSelected({ sourceId: source2.id, connectionId: source2.connections[0].id })).rejects.toThrow("Could not reuse the selected sign-in.")
})

for (const filename of ["opencode.db", "opencode-local.db"]) test(`active target ${filename} and symlink aliases cannot be offered as external sources`, () => {
  const f = fixture(filename)
  const direct = credentialReuse({ directory: f.directory, targetDirectory: f.directory, create: f.create })
  expect(direct.list().sources).toEqual([])
  const linked = mkdtempSync(path.join(os.tmpdir(), "dochaus-auth-alias-"))
  roots.push(linked)
  symlinkSync(f.file, path.join(linked, "opencode-local.db"))
  const alias = credentialReuse({ directory: linked, targetDirectory: f.directory, create: f.create })
  expect(alias.list().sources).toEqual([])
})

test("HTTP enforces native auth, origin, strict JSON, bounded body and metadata-only GET", async () => {
  const f = fixture()
  const handler = credentialReuseHandler(f.service, "synthetic-server-password")
  const auth = { Authorization: `Basic ${btoa("opencode:synthetic-server-password")}` }
  expect((await handler(new Request("http://localhost/reuse"))).status).toBe(401)
  expect((await handler(new Request("http://localhost/reuse", { headers: { ...auth, Origin: "https://foreign.invalid" } }))).status).toBe(403)
  for (const type of ["text/plain", "application/jsonp"]) expect((await handler(new Request("http://localhost/reuse", { method: "POST", headers: { ...auth, "Content-Type": type }, body: "{}" }))).status).toBe(415)
  expect((await handler(new Request("http://localhost/reuse", { method: "POST", headers: { ...auth, "Content-Type": "application/json" }, body: " ".repeat(2048) }))).status).toBe(400)
  expect((await handler(new Request("http://localhost/reuse", { method: "OPTIONS", headers: { Origin: "http://localhost:5173" } }))).status).toBe(204)
  const list = await handler(new Request("http://localhost/reuse", { headers: auth }))
  expect(list.headers.get("cache-control")).toBe("no-store")
  expect((await list.json()).sources).toHaveLength(1)
  expect(f.imported).toEqual([])
})

test("native failed insert rolls back activation and cannot log credential insert parameters", async () => {
  const { Credential } = await import("@opencode/core/credential")
  const { Database } = await import("@opencode/core/database/database")
  const { Bus } = await import("@opencode/core/bus")
  const { Global } = await import("@opencode/util/global")
  const { LayerNode } = await import("@opencode/util/effect/layer-node")
  const { Effect, Logger } = await import("effect")
  const { NodeServices } = await import("@effect/platform-node")
  const f = fixture()
  const logged: unknown[] = []
  const layer = LayerNode.compile(LayerNode.group([Credential.node, Database.node, Bus.node]), {
    replacements: [Global.node.replace(Global.layerWith({
      home: f.directory, data: f.directory, config: path.join(f.directory, "config"), state: path.join(f.directory, "state"),
      cache: path.join(f.directory, "cache"), bin: path.join(f.directory, "bin"), log: path.join(f.directory, "log"), repos: path.join(f.directory, "repos"),
    }))],
  })
  await Effect.runPromise(Effect.scoped(Effect.gen(function* () {
    const credentials = yield* Credential.Service
    const db = (yield* Database.Service).db
    const prior = yield* createReusedCredential(credentials, { integrationID: Integration.ID.make("opencode-go"), label: "Existing", value: { type: "key", key: "synthetic-existing" } })
    yield* db.run("CREATE TRIGGER reject_reuse BEFORE INSERT ON credential BEGIN SELECT RAISE(FAIL, 'synthetic-destination-error'); END")
    const result = yield* Effect.exit(createReusedCredential(credentials, { integrationID: Integration.ID.make("opencode-go"), label: "Selected", value: { type: "key", key: "synthetic-never-log-secret" } }))
    expect(result._tag).toBe("Failure")
    const row = yield* db.get<{ active: number }>(`SELECT active FROM credential WHERE id = '${prior.id}'`)
    expect(row?.active).toBe(1)
    expect(yield* credentials.list(Integration.ID.make("opencode-go"))).toHaveLength(1)
  }).pipe(Effect.provide(layer), Effect.provide(NodeServices.layer), Effect.provide(Logger.layer([Logger.make(entry => { logged.push(entry.message) })])))))
  expect(JSON.stringify(logged)).not.toContain("synthetic-never-log-secret")
  expect(JSON.stringify(logged)).not.toContain("synthetic-destination-error")
})

test("read-only discovery sees committed WAL credentials without changing source files", async () => {
  const f = fixture()
  const writer = new Database(f.file)
  try {
    writer.exec("PRAGMA journal_mode=WAL")
    writer.query("INSERT INTO credential VALUES (?, ?, ?, ?)").run("wal-key", "openai", "Recent", JSON.stringify(f.key))
    const files = readdirSync(f.directory)
    const before = files.map(file => readFileSync(path.join(f.directory, file)))
    const source = f.service.list().sources[0]
    expect(source.connections).toHaveLength(3)
    await f.service.importSelected({ sourceId: source.id, connectionId: source.connections.find(connection => connection.integrationID === "openai")!.id })
    expect(f.imported).toHaveLength(1)
    expect(readdirSync(f.directory)).toEqual(files)
    // SQLite may update the existing SHM read mark for WAL coordination;
    // the credential database and committed WAL remain byte-for-byte intact.
    files.forEach((file, index) => { if (!file.endsWith("-shm")) expect(readFileSync(path.join(f.directory, file))).toEqual(before[index]) })
  } finally { writer.close() }
})
