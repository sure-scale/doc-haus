import { Hono } from "hono"
import { existsSync } from "node:fs"
import path from "node:path"
import { listKeyDates, type KeyDateKind, type KeyDateRow } from "../../../dochaus/lib/key-dates"

// Read-only view of the matter's key-dates register. The rows are written by
// the record-key-date tool during reviews; this route only shapes them for the
// web register, so a later surface cannot silently mutate a date by reading it.
// The register lives in the matter's legal.db, beside the other dochaus tables.

export type KeyDateRegisterRow = {
  id: number
  document: string
  kind: KeyDateKind
  date: string | null
  formula: string | null
  noticeDays: number | null
  actBy: string | null
  party: string | null
  quote: string
  reason: string
  author: string
  recordedAt: number
}

export function keyDateRegisterRow(row: KeyDateRow): KeyDateRegisterRow {
  return {
    id: row.id,
    document: row.doc_name,
    kind: row.kind,
    date: row.date_iso,
    formula: row.date_formula,
    noticeDays: row.notice_days,
    actBy: row.cancel_by_iso,
    party: row.party,
    quote: row.quote,
    reason: row.reason,
    author: row.author,
    recordedAt: row.created_at,
  }
}

export function keyDateRoutes(resolveMatter: (id: string) => string) {
  const app = new Hono()
  app.use("/matters/:id/key-dates", async (c, next) => {
    const id = c.req.param("id") ?? ""
    if (!/^[a-z0-9][a-z0-9-]*$/.test(id)) return c.json({ error: "Invalid matter id." }, 400)
    const dir = resolveMatter(id)
    if (!existsSync(path.join(dir, "matter.json"))) return c.json({ error: "Matter not found." }, 404)
    await next()
  })
  app.get("/matters/:id/key-dates", (c) =>
    c.json({ rows: listKeyDates(resolveMatter(c.req.param("id"))).map(keyDateRegisterRow) }),
  )
  return app
}
