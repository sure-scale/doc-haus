import { Database } from "bun:sqlite"
import { mkdirSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { workspaceRoot } from "./matter"

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
const toClient = ({ reference, ...row }: Row): Client => ({ ...row, ...(reference ? { reference } : {}) })

export function listClients(): Client[] {
  const db = open()
  try {
    return db.query<Row, []>("SELECT id, name, reference, created_at FROM clients ORDER BY name COLLATE NOCASE").all().map(toClient)
  } finally {
    db.close()
  }
}

// Adding a name that is already a client returns that client unchanged.
export function addClient(input: { name?: unknown; reference?: unknown }): { client: Client; created: boolean } {
  if (typeof input.name !== "string" || !line(input.name)) throw new ClientValidationError("Name the client to add it.")
  if (input.reference !== undefined && typeof input.reference !== "string") throw new ClientValidationError("The client number must be text.")
  const name = line(input.name)
  if (name.length > NAME_MAX_LENGTH)
    throw new ClientValidationError(`A client name can be at most ${NAME_MAX_LENGTH} characters.`)
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
  app.delete("/clients/:id", (c) => {
    const id = Number(c.req.param("id"))
    if (!Number.isInteger(id)) return c.json({ error: "invalid client id" }, 400)
    removeClient(id)
    return c.json({ ok: true })
  })
  return app
}
