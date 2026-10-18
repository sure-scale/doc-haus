import { Hono } from "hono"
import path from "node:path"
import { mkdirSync } from "node:fs"
import { listPlaybooks, workspaceRoot } from "./matter"

// Source uploads live outside the discoverable skills directory. They only
// become selectable playbooks after the lawyer submits the builder's import.
export function playbookLibraryRoutes() {
  const app = new Hono()
  const dir = path.join(workspaceRoot(), ".playbook-library")
  app.get("/playbooks/library", (c) => {
    mkdirSync(dir, { recursive: true })
    return c.json({ dir, playbooks: listPlaybooks() })
  })
  app.post("/playbooks/source", async (c) => {
    const body = await c.req.parseBody()
    const file = body.file
    if (!(file instanceof File)) return c.json({ error: "Choose a playbook file to import" }, 400)
    const name = path.basename(file.name)
    const ext = path.extname(name).toLowerCase()
    if (![".docx", ".md", ".txt"].includes(ext))
      return c.json({ error: "Playbooks must be .docx, .md or .txt files" }, 400)
    try {
      const buffer = Buffer.from(await file.arrayBuffer())
      const text = ext === ".docx" ? await extractSource(name, buffer) : buffer.toString("utf8")
      if (!text.trim()) return c.json({ error: "This playbook has no readable text" }, 400)
      const source = `imports/${crypto.randomUUID()}.txt`
      await Bun.write(path.join(dir, source), text)
      return c.json({ name, source }, 201)
    } catch {
      return c.json({ error: "Could not read this playbook. Check the file and try again." }, 400)
    }
  })
  return app
}

async function extractSource(name: string, buffer: Buffer) {
  const { extractDocumentText } = await import("./ingest")
  return extractDocumentText(name, buffer)
}
