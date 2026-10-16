import { Hono } from "hono"
import { openDb } from "./db"

// Display labels are user metadata, independent of retrieval roles and drafting gates.
export function documentLabelRoutes(resolveMatter: (id: string) => string) {
  const app = new Hono()
  app.patch("/matters/:id/documents/label", async (c) => {
    const input = await c.req.json<{ name?: unknown; label?: unknown }>().catch(() => null)
    if (!input || typeof input.name !== "string" || !input.name ||
      typeof input.label !== "string" || !["draft", "redline", "notes", "final"].includes(input.label))
      return c.json({ error: "Choose Draft, Redline, Notes, or Final for an existing document." }, 400)
    const db = openDb(resolveMatter(c.req.param("id")))
    try {
      const result = db.run("UPDATE documents SET label = ? WHERE name = ?", [input.label, input.name])
      if (!result.changes) return c.json({ error: "Document not found." }, 404)
      return c.json({ name: input.name, label: input.label })
    } finally {
      db.close()
    }
  })
  return app
}
