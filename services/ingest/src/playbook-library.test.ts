import { afterAll, expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, rmSync, symlinkSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"

const root = mkdtempSync(path.join(tmpdir(), "dochaus-playbook-library-"))
process.env.WORKSPACE_ROOT = root
const { playbookLibraryRoutes } = await import("./playbook-library")
const matter = await import("./matter")
const app = playbookLibraryRoutes()
afterAll(() => rmSync(root, { recursive: true, force: true }))

async function upload(name: string, content: BlobPart) {
  const body = new FormData()
  body.append("file", new File([content], name))
  return app.request("/playbooks/source", { method: "POST", body })
}

test("source uploads preserve text and remain outside the selectable playbook library", async () => {
  const content = "## Payment\nPreferred: thirty days.\n```approved\nPayment is due within thirty days.\n```"
  const res = await upload("Firm positions.md", content)
  expect(res.status).toBe(201)
  const source = await res.json()
  expect(source.name).toBe("Firm positions.md")
  expect(readFileSync(path.join(root, ".playbook-library", source.source), "utf8")).toBe(content)
  const library = await (await app.request("/playbooks/library")).json()
  expect(library.dir).toBe(path.join(root, ".playbook-library"))
  expect(library.playbooks).toEqual(matter.listPlaybooks())
  expect(library.playbooks.some((p: { content: string }) => p.content === content)).toBe(false)
  const second = await (await upload("Firm positions.md", "Revised positions")).json()
  expect(second.source).not.toBe(source.source)
  expect(readFileSync(path.join(root, ".playbook-library", source.source), "utf8")).toBe(content)
})

test("DOCX sources use the document extractor and retain operative wording", async () => {
  const clause = "Payment shall be made within thirty days of receipt."
  const bytes = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(clause)] }] }))
  const res = await upload("Firm policy.docx", new Uint8Array(bytes))
  expect(res.status).toBe(201)
  const source = await res.json()
  expect(readFileSync(path.join(root, ".playbook-library", source.source), "utf8").trim()).toBe(clause)
})

test("library chat sources are readable in each assistant directory and are not library entries", async () => {
  for (const [kind, directory] of [["playbooks", ".playbook-library"], ["templates", ".templates"], ["skills", ".skills"], ["workflows", ".workflows"], ["agents", ".agents"]]) {
    const body = new FormData()
    body.append("file", new File(["Confidential source text"], "guide.markdown"))
    const res = await app.request(`/library/${kind}/sources`, { method: "POST", body })
    expect(res.status).toBe(201)
    const source = await res.json()
    expect(source.name).toBe("guide.markdown")
    expect(readFileSync(path.join(root, directory!, source.source), "utf8")).toBe("Confidential source text")
    expect(source.source).toMatch(/^\.chat-attachments\/[a-f0-9-]+\/text\.txt$/)
    expect(readFileSync(path.join(root, directory!, source.original), "utf8")).toBe("Confidential source text")
    expect((await app.request(`/library/${kind}/sources/${source.id}`, { method: "DELETE" })).status).toBe(200)
    expect(existsSync(path.join(root, directory!, source.original))).toBe(false)
    expect(existsSync(path.join(root, directory!, source.source))).toBe(false)
  }
  const body = new FormData()
  body.append("file", new File(["bad"], "bad.exe"))
  expect((await app.request("/library/skills/sources", { method: "POST", body })).status).toBe(400)
  expect((await app.request("/library/other/sources", { method: "POST", body })).status).toBe(404)
})

test("library chat extracts PDF text for the assistant", async () => {
  const body = new FormData()
  body.append("file", new File([readFileSync(path.join(import.meta.dir, "fixtures/synthetic-text-layer.pdf"))], "guide.pdf"))
  const res = await app.request("/library/playbooks/sources", { method: "POST", body })
  expect(res.status).toBe(201)
  const source = await res.json()
  expect(readFileSync(path.join(root, ".playbook-library", source.source), "utf8").trim().length).toBeGreaterThan(0)
  expect(readFileSync(path.join(root, ".playbook-library", source.original))).toEqual(readFileSync(path.join(import.meta.dir, "fixtures/synthetic-text-layer.pdf")))
  const imported = await upload("guide.pdf", new Uint8Array(readFileSync(path.join(import.meta.dir, "fixtures/synthetic-text-layer.pdf"))))
  expect(imported.status).toBe(201)
  const importedSource = await imported.json()
  expect(readFileSync(path.join(root, ".playbook-library", importedSource.source), "utf8").trim().length).toBeGreaterThan(0)
})

test("chat attachment preserves original Word bytes and rejects invalid removal ids", async () => {
  const bytes = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Original text")] }] }))
  const body = new FormData()
  body.append("file", new File([new Uint8Array(bytes)], "source.docx"))
  const res = await app.request("/library/templates/sources", { method: "POST", body })
  const source = await res.json()
  expect(Array.from(readFileSync(path.join(root, ".templates", source.original)))).toEqual(Array.from(bytes))
  expect((await app.request("/library/templates/sources/not-a-uuid", { method: "DELETE" })).status).toBe(400)
  expect(existsSync(path.join(root, ".templates", source.original))).toBe(true)
})

test("attachment staging refuses a symlinked attachment directory", async () => {
  const { stageChatAttachment } = await import("./chat-attachments")
  const directory = mkdtempSync(path.join(root, "symlink-check-"))
  symlinkSync(root, path.join(directory, ".chat-attachments"), "dir")
  await expect(stageChatAttachment(directory, new File(["private"], "guide.txt"))).rejects.toThrow("Invalid attachment directory")
})

test("missing, unsupported, empty and broken files return actionable failures", async () => {
  expect((await app.request("/playbooks/source", { method: "POST", body: new FormData() })).status).toBe(400)
  for (const [name, content] of [
    ["guide.exe", "text"],
    ["guide.txt", "   "],
    ["broken.docx", "not a word document"],
  ]) {
    const res = await upload(name, content)
    expect(res.status).toBe(400)
    expect((await res.json()).error).toBeTruthy()
  }
})

test("edits retain bindings and metadata, and deletion requires unbinding", async () => {
  const name = "playbook-library-check"
  await Bun.write(
    path.join(root, ".playbooks", name, "SKILL.md"),
    `---\nname: ${name}\ndescription: "Original"\nversion: "2"\nlast-reviewed: "2026-10-01"\n---\n\n## Payment\nOriginal body`,
  )
  const bound = matter.createMatter("Playbook test matter", "PB-1", ["EW"], name)
  const updated = matter.updatePlaybook(name, { description: "Updated", content: "## Payment\nRevised body" })
  expect(updated.version).toBe("2")
  expect(updated.last_reviewed).toBe("2026-10-01")
  expect(updated.matters).toEqual([bound.title])
  expect(() => matter.deletePlaybook(name)).toThrow("Unbind it from each matter before deleting")
  expect(matter.getMatter(bound.id).playbook).toBe(name)
  matter.renameMatter(bound.id, bound.title, bound.reference, bound.jurisdictions, "")
  matter.deletePlaybook(name)
  expect(matter.listPlaybooks().some((p) => p.name === name)).toBe(false)
})
