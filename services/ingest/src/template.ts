import { existsSync, mkdirSync, readdirSync, copyFileSync, readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { docxodus } from "./docxodus"
import { WORKSPACE_ROOT } from "./matter"

// The global, user-managed template library: drafting templates the firm can
// add, remove, and create, shared across every matter. It lives at
// WORKSPACE_ROOT/.templates — invisible to listMatters() (no matter.json) and to
// matterDir() (the matter-id regex rejects a leading dot), so it never collides
// with a matter directory. Ingest is the single writer over WORKSPACE_ROOT, so
// the library is owned here; the dochaus tools reach it over HTTP.

export const TEMPLATES_DIR = path.join(WORKSPACE_ROOT, ".templates")

// One-line descriptions of what each template is for, keyed by file name. Kept in
// an ingest-owned manifest beside the .docx files rather than inside them, so a
// description can be edited without touching the document bytes and so listing
// stays a cheap file read alongside the placeholder scan.
const MANIFEST = path.join(TEMPLATES_DIR, "manifest.json")

// A template name arrives from a URL query or multipart field. basename() is the
// single containment guard keeping the lookup inside the library directory,
// matching the document routes' guard.
export function templatePath(name: string) {
  return path.join(TEMPLATES_DIR, path.basename(name))
}

// Seed the library from repo demo/templates when the sample matter is added. Only
// seedSampleMatter (seed.ts) calls this — an install without the sample starts
// with an empty library. The directory's own existence is the first-run marker:
// idempotent across restarts, and it never resurrects a template the user has
// deleted (a present dir means we already seeded). _base.docx stays repo-internal
// under dochaus/templates — it is the from-scratch drafting seed, never a
// user-visible template.
export function seedTemplates() {
  if (existsSync(TEMPLATES_DIR)) return
  mkdirSync(TEMPLATES_DIR, { recursive: true })
  const demoTemplates = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "demo", "templates")
  // Curated optional demo templates may be absent from a source distribution.
  if (!existsSync(demoTemplates)) return
  readdirSync(demoTemplates)
    .filter((name) => name.endsWith(".docx"))
    .forEach((name) => copyFileSync(path.join(demoTemplates, name), templatePath(name)))
}

export async function listTemplates() {
  mkdirSync(TEMPLATES_DIR, { recursive: true })
  const names = readdirSync(TEMPLATES_DIR).filter((n) => n.endsWith(".docx"))
  const manifest = readManifest()
  const dx = await docxodus()
  return Promise.all(
    names.map(async (name) => {
      const session = dx.openDocxSession(await Bun.file(templatePath(name)).bytes(), {})
      const placeholders = session.findPlaceholders().map((p) => ({ text: p.match.text, kind: p.kind, hint: p.hint }))
      session.close()
      return { name, description: manifest[name] ?? "", placeholders }
    }),
  )
}

export function setTemplateDescription(name: string, description: string) {
  const key = path.basename(name)
  writeManifest({ ...readManifest(), [key]: description })
  return { name: key, description }
}

export function removeTemplateDescription(name: string) {
  const manifest = readManifest()
  delete manifest[path.basename(name)]
  writeManifest(manifest)
}

function readManifest(): Record<string, string> {
  if (!existsSync(MANIFEST)) return {}
  return JSON.parse(readFileSync(MANIFEST, "utf8"))
}

function writeManifest(manifest: Record<string, string>) {
  mkdirSync(TEMPLATES_DIR, { recursive: true })
  writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2))
}
