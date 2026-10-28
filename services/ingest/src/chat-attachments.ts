import path from "node:path"
import { lstat, mkdir, rm, writeFile } from "node:fs/promises"

async function attachmentRoot(dir: string) {
  const root = path.join(dir, ".chat-attachments")
  await mkdir(root, { recursive: true, mode: 0o700 })
  const stat = await lstat(root)
  if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error("Invalid attachment directory")
  return root
}

// Each attachment has its own generated directory, so removal cannot affect a
// different attachment and the original remains available for document work.
export async function stageChatAttachment(dir: string, file: File) {
  const name = path.basename(file.name)
  const ext = path.extname(name).toLowerCase()
  if (![".docx", ".pdf", ".md", ".markdown", ".txt"].includes(ext))
    throw new Error("Attach a PDF, Word, Markdown or text file")
  const buffer = Buffer.from(await file.arrayBuffer())
  let text: string
  try {
    if (ext === ".docx" || ext === ".pdf") {
      const { extractDocumentText } = await import("./ingest")
      text = await extractDocumentText(name, buffer)
    } else text = buffer.toString("utf8")
  } catch {
    throw new Error("Could not read this file. Check the file and try again.")
  }
  if (!text.trim()) throw new Error("This file has no readable text")
  const id = crypto.randomUUID()
  const relativeDir = `.chat-attachments/${id}`
  const root = await attachmentRoot(dir)
  const target = path.join(root, id)
  const source = `${relativeDir}/text.txt`
  const original = `${relativeDir}/original${ext}`
  try {
    await mkdir(target, { mode: 0o700 })
    await writeFile(path.join(dir, original), buffer, { mode: 0o600, flag: "wx" })
    await writeFile(path.join(dir, source), text, { mode: 0o600, flag: "wx" })
  } catch (error) {
    await rm(target, { recursive: true, force: true })
    throw error
  }
  return { id, name, source, original }
}

export async function removeChatAttachment(dir: string, id: string) {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(id))
    throw new Error("Invalid attachment")
  const root = await attachmentRoot(dir)
  await rm(path.join(root, id), { recursive: true, force: true })
}
