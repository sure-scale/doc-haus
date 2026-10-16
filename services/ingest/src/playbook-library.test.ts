import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
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
