import { Hono } from "hono"
import { copyWordDocument, copyWordDocumentName, DeliverableConflictError } from "./deliverable"

export function documentCopyRoutes(resolveMatter: (id: string) => string) {
  return new Hono().post("/matters/:id/documents/copy", async (c) => {
    const body = await c.req.json().catch(() => null) as { source?: unknown; name?: unknown } | null
    if (typeof body?.source !== "string" || typeof body.name !== "string" ||
      !copyWordDocumentName(body.source) || !copyWordDocumentName(body.name))
      return c.json({ error: "Use source and name Word filenames within this matter, without paths." }, 400)
    try {
      return c.json(await copyWordDocument(resolveMatter(c.req.param("id")), body.source, body.name))
    } catch (error) {
      if (error instanceof DeliverableConflictError) return c.json({ error: error.message }, 409)
      throw error
    }
  })
}
