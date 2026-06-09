import { Hono } from "hono"
import { cors } from "hono/cors"
import { openDb, listDocuments, deleteDocument } from "./db"
import { ingestDocx } from "./ingest"
import { listMatters, createMatter, getMatter, deleteMatter, matterDir } from "./matter"
import { existsSync, rmSync } from "node:fs"
import path from "node:path"

const app = new Hono()
app.use("*", cors())

app.get("/matters", (c) => c.json(listMatters()))

app.post("/matters", async (c) => {
  const { title, reference } = await c.req.json<{ title: string; reference?: string }>()
  return c.json(createMatter(title, reference))
})

app.delete("/matters/:id", (c) => {
  deleteMatter(c.req.param("id"))
  return c.json({ ok: true })
})

app.get("/matters/:id", (c) => {
  const matter = getMatter(c.req.param("id"))
  const dbPath = path.join(matter.dir, ".dochaus", "legal.db")
  const documents = existsSync(dbPath) ? listDocuments(openDb(matter.dir)) : []
  return c.json({ ...matter, documents })
})

app.post("/matters/:id/documents", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const body = await c.req.parseBody()
  const file = body["file"] as File
  const buffer = Buffer.from(await file.arrayBuffer())
  const result = await ingestDocx(dir, file.name, buffer)
  return c.json(result)
})

// Remove a document: delete the source .docx and its index rows. basename()
// keeps the lookup inside the matter directory, matching the content route.
app.delete("/matters/:id/documents", (c) => {
  const dir = matterDir(c.req.param("id"))
  const file = path.join(dir, path.basename(c.req.query("name") ?? ""))
  rmSync(file, { force: true })
  if (existsSync(path.join(dir, ".dochaus", "legal.db"))) deleteDocument(openDb(dir), file)
  return c.json({ ok: true })
})

// Serve a matter's .docx bytes so the web app can render the redline in-browser
// (the viewer converts to HTML client-side via WASM — the file is never uploaded
// anywhere). basename() keeps the lookup inside the matter directory.
app.get("/matters/:id/documents/content", async (c) => {
  const file = path.join(matterDir(c.req.param("id")), path.basename(c.req.query("name") ?? ""))
  if (!existsSync(file)) return c.notFound()
  return new Response(Bun.file(file).stream(), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `inline; filename="${path.basename(file)}"`,
    },
  })
})

const port = Number(process.env.INGEST_PORT ?? 4500)
console.log(`doc.haus ingest service listening on http://localhost:${port}`)
export default { port, fetch: app.fetch }
