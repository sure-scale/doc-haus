import { Hono } from "hono"
import { cors } from "hono/cors"
import { openDb, listDocuments } from "./db"
import { ingestDocx } from "./ingest"
import { listMatters, createMatter, getMatter, matterDir } from "./matter"
import { existsSync } from "node:fs"
import path from "node:path"

const app = new Hono()
app.use("*", cors())

app.get("/matters", (c) => c.json(listMatters()))

app.post("/matters", async (c) => {
  const { title } = await c.req.json<{ title: string }>()
  return c.json(createMatter(title))
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
