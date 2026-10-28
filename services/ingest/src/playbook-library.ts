import { Hono } from "hono"
import path from "node:path"
import { mkdirSync } from "node:fs"
import { listPlaybooks, workspaceRoot } from "./matter"
import { removeChatAttachment, stageChatAttachment } from "./chat-attachments"

const chatDirs: Record<string, string> = { playbooks: ".playbook-library", templates: ".templates", skills: ".skills", workflows: ".workflows", agents: ".agents" }
function chatDir(kind: string) {
  return Object.hasOwn(chatDirs, kind) ? path.join(workspaceRoot(), chatDirs[kind]!) : undefined
}

// Source uploads live outside the discoverable skills directory. They only
// become selectable playbooks after the lawyer submits the builder's import.
export function playbookLibraryRoutes() {
  const app = new Hono()
  // Read per request so the directory follows WORKSPACE_ROOT.
  const libraryDir = () => path.join(workspaceRoot(), ".playbook-library")
  app.get("/playbooks/library", (c) => {
    const dir = libraryDir()
    mkdirSync(dir, { recursive: true })
    return c.json({ dir, playbooks: listPlaybooks() })
  })
  app.post("/playbooks/source", async (c) => {
    const body = await c.req.parseBody()
    const file = body.file
    if (!(file instanceof File)) return c.json({ error: "Choose a playbook file to import" }, 400)
    const dir = libraryDir()
    const name = path.basename(file.name)
    const ext = path.extname(name).toLowerCase()
    if (![".docx", ".pdf", ".md", ".markdown", ".txt"].includes(ext))
      return c.json({ error: "Playbooks must be PDF, Word, Markdown or text files" }, 400)
    try {
      const buffer = Buffer.from(await file.arrayBuffer())
      const text = ext === ".docx" || ext === ".pdf" ? await extractSource(name, buffer) : buffer.toString("utf8")
      if (!text.trim()) return c.json({ error: "This playbook has no readable text" }, 400)
      const source = `imports/${crypto.randomUUID()}.txt`
      await Bun.write(path.join(dir, source), text)
      return c.json({ name, source }, 201)
    } catch {
      return c.json({ error: "Could not read this playbook. Check the file and try again." }, 400)
    }
  })
  // Library chat sources are staged outside the discoverable library entries.
  // Keep them inside the assistant's working directory for subsequent turns.
  app.post("/library/:kind/sources", async (c) => {
    const dir = chatDir(c.req.param("kind"))
    if (!dir) return c.json({ error: "Unknown library" }, 404)
    const file = (await c.req.parseBody()).file
    if (!(file instanceof File)) return c.json({ error: "Choose a file to attach" }, 400)
    try {
      return c.json(await stageChatAttachment(dir, file), 201)
    } catch (error) {
      return c.json({ error: error instanceof Error ? error.message : "Could not attach this file" }, 400)
    }
  })
  app.delete("/library/:kind/sources/:id", async (c) => {
    const dir = chatDir(c.req.param("kind"))
    if (!dir) return c.json({ error: "Unknown library" }, 404)
    try {
      await removeChatAttachment(dir, c.req.param("id"))
      return c.json({ ok: true })
    } catch {
      return c.json({ error: "Could not remove attachment" }, 400)
    }
  })
  return app
}

async function extractSource(name: string, buffer: Buffer) {
  const { extractDocumentText } = await import("./ingest")
  return extractDocumentText(name, buffer)
}
