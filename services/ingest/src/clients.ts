import { Database } from "bun:sqlite"
import { mkdirSync } from "node:fs"
import path from "node:path"
import { Hono, type Context } from "hono"
import { listMatters, setMatterClient, workspaceRoot } from "./matter"

// Clients the firm added on their own, before or without a matter. The web
// Clients list joins these with the Client named on each matter, so a matter's
// client needs no record here. One small workspace database, outside any
// matter directory; the per-matter legal.db (db.ts) is scoped to its matter.

export type Client = { id: number; name: string; reference?: string; created_at: number }

// Matches the matter Client cap (matter.ts), so any stored name can be a matter's client.
const NAME_MAX_LENGTH = 200

class ClientValidationError extends Error {}

function line(value: string) {
  return value.replace(/\p{Cc}/gu, " ").normalize("NFC").replace(/\s+/g, " ").trim()
}

// The same identity the web groups matter clients by (apps/web/src/lib/clients.ts):
// case, spacing, Unicode composition and trailing punctuation do not make a new client.
export function clientKey(name: string) {
  return line(name).replace(/[.,;:]+$/, "").trimEnd().toLowerCase()
}

function open() {
  const dir = path.join(workspaceRoot(), ".clients")
  mkdirSync(dir, { recursive: true })
  const db = new Database(path.join(dir, "clients.db"))
  db.run("CREATE TABLE IF NOT EXISTS clients (id INTEGER PRIMARY KEY, key TEXT NOT NULL UNIQUE, name TEXT NOT NULL, reference TEXT, created_at INTEGER NOT NULL)")
  return db
}

type Row = { id: number; name: string; reference: string | null; created_at: number }
const toClient = ({ id, name, reference, created_at }: Row): Client => ({ id, name, ...(reference ? { reference } : {}), created_at })

export function listClients(): Client[] {
  const db = open()
  try {
    return db.query<Row, []>("SELECT id, name, reference, created_at FROM clients ORDER BY name COLLATE NOCASE").all().map(toClient)
  } finally {
    db.close()
  }
}

function validName(value: unknown, missing: string) {
  if (typeof value !== "string" || !line(value)) throw new ClientValidationError(missing)
  const name = line(value)
  if (name.length > NAME_MAX_LENGTH)
    throw new ClientValidationError(`A client name can be at most ${NAME_MAX_LENGTH} characters.`)
  return name
}

// Adding a name that is already a client returns that client unchanged.
export function addClient(input: { name?: unknown; reference?: unknown }): { client: Client; created: boolean } {
  const name = validName(input.name, "Name the client to add it.")
  if (input.reference !== undefined && typeof input.reference !== "string") throw new ClientValidationError("The client number must be text.")
  const key = clientKey(name)
  const db = open()
  try {
    const inserted = db
      .query("INSERT INTO clients (key, name, reference, created_at) VALUES (?, ?, ?, ?) ON CONFLICT(key) DO NOTHING")
      .run(key, name, line(input.reference ?? "") || null, Date.now())
    const row = db.query<Row, [string]>("SELECT id, name, reference, created_at FROM clients WHERE key = ?").get(key)!
    return { client: toClient(row), created: inserted.changes > 0 }
  } finally {
    db.close()
  }
}

export class ClientNotFoundError extends Error {}


// Edits a client's name and number. The source is a stored client (id) or a
// client named only on matters (from), which gains a stored record here. Every
// matter whose Client is the old client (same key) is renamed with it, so the
// client never splits in two. A name that is already another client merges into
// that client: its stored name and number are kept (a blank number takes the
// edited one) and the old record goes.
export function editClient(
  source: { id: number } | { from: unknown },
  input: { name?: unknown; reference?: unknown },
): { client: Client; matters: number; merged: boolean } {
  const name = validName(input.name, "Name the client to save it.")
  if (input.reference !== undefined && typeof input.reference !== "string") throw new ClientValidationError("The client number must be text.")
  const key = clientKey(name)
  const fromKey = "from" in source ? clientKey(validName(source.from, "Name the client to edit.")) : undefined
  const db = open()
  let oldKey: string
  let result: { client: Client; merged: boolean }
  try {
    const select = "SELECT id, key, name, reference, created_at FROM clients WHERE"
    const byKey = db.query<Row & { key: string }, [string]>(`${select} key = ?`)
    const current = "id" in source ? db.query<Row & { key: string }, [number]>(`${select} id = ?`).get(source.id) : byKey.get(fromKey!)
    if (!current && "id" in source) throw new ClientNotFoundError("That client no longer exists.")
    oldKey = current?.key ?? fromKey!
    const reference = input.reference === undefined ? current?.reference ?? null : line(input.reference) || null
    const target = key === oldKey ? null : byKey.get(key)
    result = db.transaction(() => {
      if (target) {
        if (current) db.query("DELETE FROM clients WHERE id = ?").run(current.id)
        if (!target.reference && reference) db.query("UPDATE clients SET reference = ? WHERE id = ?").run(reference, target.id)
        return { client: toClient({ ...target, reference: target.reference || reference }), merged: true }
      }
      if (current) {
        db.query("UPDATE clients SET key = ?, name = ?, reference = ? WHERE id = ?").run(key, name, reference, current.id)
        return { client: toClient({ ...current, name, reference }), merged: false }
      }
      const created_at = Date.now()
      const { lastInsertRowid } = db.query("INSERT INTO clients (key, name, reference, created_at) VALUES (?, ?, ?, ?)").run(key, name, reference, created_at)
      return { client: toClient({ id: Number(lastInsertRowid), name, reference, created_at }), merged: false }
    })()
  } finally {
    db.close()
  }
  // Matters are renamed after the record: if one fails, editing the old name
  // again (now a client only on its remaining matters) finishes the move.
  let matters = 0
  for (const matter of listMatters()) {
    if (!matter.client || clientKey(matter.client) !== oldKey) continue
    if (matter.client !== result.client.name) setMatterClient(matter.id, result.client.name)
    matters++
  }
  return { ...result, matters }
}

export function removeClient(id: number) {
  const db = open()
  try {
    db.query("DELETE FROM clients WHERE id = ?").run(id)
  } finally {
    db.close()
  }
}

export function clientRoutes() {
  const app = new Hono()
  app.get("/clients", (c) => c.json(listClients()))
  // 201 for a new client, 200 with the existing one for a name already added.
  app.post("/clients", async (c) => {
    const body = await c.req.json().catch(() => undefined)
    try {
      const { client, created } = addClient(body && typeof body === "object" ? body : {})
      return c.json(client, created ? 201 : 200)
    } catch (error) {
      if (error instanceof ClientValidationError) return c.json({ error: error.message }, 400)
      throw error
    }
  })
  // Edit a stored client by id, or one named only on matters by its current
  // name (rename, body.from); both return { client, matters, merged }.
  async function edit(c: Context, source: (body: Record<string, unknown>) => { id: number } | { from: unknown }) {
    const parsed = await c.req.json().catch(() => undefined)
    const body: Record<string, unknown> = parsed && typeof parsed === "object" ? parsed : {}
    try {
      return c.json(editClient(source(body), body))
    } catch (error) {
      if (error instanceof ClientValidationError) return c.json({ error: error.message }, 400)
      if (error instanceof ClientNotFoundError) return c.json({ error: error.message }, 404)
      throw error
    }
  }
  app.patch("/clients/:id", (c) => {
    const id = Number(c.req.param("id"))
    if (!Number.isInteger(id)) return c.json({ error: "invalid client id" }, 400)
    return edit(c, () => ({ id }))
  })
  app.post("/clients/rename", (c) => edit(c, (body) => ({ from: body.from })))
  app.delete("/clients/:id", (c) => {
    const id = Number(c.req.param("id"))
    if (!Number.isInteger(id)) return c.json({ error: "invalid client id" }, 400)
    removeClient(id)
    return c.json({ ok: true })
  })
  return app
}
