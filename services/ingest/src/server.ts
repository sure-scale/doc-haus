import { Hono } from "hono"
import { cors } from "hono/cors"
import {
  openDb,
  listDocuments,
  deleteDocument,
  listPendingRedlines,
  getRedline,
  setRedlineStatus,
  pendingRedlineCounts,
  getInjectionReport,
} from "./db"
import { ingestDocument, extractDocumentText } from "./ingest"
import { pdfToDocx } from "./convert"
import { buildRedlined, bake } from "./redline"
import { listMatters, createMatter, getMatter, renameMatter, deleteMatter, matterDir, listJurisdictions, listPlaybooks, updatePlaybook, deletePlaybook, PlaybookError, WORKSPACE_ROOT } from "./matter"
import { listTemplates, templatePath, setTemplateDescription, removeTemplateDescription, TEMPLATES_DIR } from "./template"
import { listWorkflows, createWorkflow, updateWorkflow, deleteWorkflow, NAME_RE, WORKFLOWS_DIR } from "./workflow"
import { listSkills, createSkill, updateSkill, deleteSkill, importSkill, setSkillEnabled, SKILLS_DIR } from "./skill"
import { listAgents, createAgent, updateAgent, deleteAgent, setAgentEnabled, AGENTS_DIR } from "./agent"
import { docxodus } from "./docxodus"
import { listGcpProjects, listAwsProfiles, probeVertex } from "./host"
import { readGrid, writeGrid, type Grid } from "./grid"
import { readDraftingPreferences, writeDraftingPreferences } from "./preferences"
import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"

const app = new Hono()

// Mirror opencode's CORS posture (packages/opencode/src/server/cors.ts): allow
// only localhost/127.0.0.1 origins on any port, not a wildcard. The web app runs
// on localhost; a wildcard would let any site in a user's browser hit this
// service. Non-browser callers (curl, the opencode server) send no Origin and are
// unaffected.
app.use(
  "*",
  cors({ origin: (origin) => (/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin) ? origin : null) }),
)

// What the host's cloud sign-ins can see, for the provider-setup comboboxes in
// the web settings (GCP projects via gcloud ADC, AWS profiles via ~/.aws).
// Empty lists when the host has no such sign-in (e.g. gcloud not installed).
app.get("/host/gcp-projects", async (c) => c.json({ projects: await listGcpProjects().catch(() => []) }))

app.get("/host/aws-profiles", async (c) => c.json({ profiles: await listAwsProfiles() }))

// Verify a Vertex project/location with one real publisher-model call using the
// host's ADC. The web's provider gate calls this instead of prompting through
// the engine, whose first-touch cold boot can outlast any probe timeout.
app.post("/host/probe-vertex", async (c) =>
  c.json(
    await probeVertex(await c.req.json()).catch((e) => ({
      ok: false,
      error: e instanceof Error ? e.message : String(e),
    })),
  ),
)

app.get("/matters", (c) => c.json(listMatters()))

// Firm-wide drafting preferences (Settings → Drafting). Saving re-renders the
// standing-instructions markdown the engine reads on every turn (see
// preferences.ts), so a change applies from the assistant's next reply.
app.get("/preferences", (c) => c.json(readDraftingPreferences()))

app.put("/preferences", async (c) => c.json(writeDraftingPreferences(await c.req.json())))

// The jurisdiction packs a matter can be assigned, read from the dochaus config
// layer. The web app populates its matter-creation dropdown from this (issue #18).
app.get("/jurisdictions", (c) => c.json(listJurisdictions()))

// The playbook skills a matter can be reviewed against: the repo-shipped ones plus
// any the firm imported into WORKSPACE_ROOT/.playbooks. The web app offers these as
// the matter's playbook choice; create-playbook reads it to dedupe before writing.
app.get("/playbooks", (c) => c.json(listPlaybooks()))

// Import a firm playbook skill into WORKSPACE_ROOT/.playbooks. Ingest is the single
// writer over WORKSPACE_ROOT (same rationale as the .templates library), so the
// dochaus create-playbook tool POSTs here rather than writing the file itself. The
// frontmatter (name + quoted description) is added here; the body is supplied raw.
app.post("/playbooks", async (c) => {
  const { name, description, content } = await c.req.json<{ name: string; description: string; content: string }>()
  if (!name.startsWith("playbook-")) return c.json({ error: 'Playbook name must start with "playbook-"' }, 400)
  const skill = path.basename(name)
  // Same slug rule skill.ts enforces — the name goes onto an unquoted YAML
  // `name:` line below, so only lowercase/hyphen slugs are safe to write.
  if (!NAME_RE.test(skill)) return c.json({ error: `"${skill}" is not a valid playbook name (lowercase, hyphenated)` }, 400)
  if (listPlaybooks().some((p) => p.name === skill)) return c.json({ error: `playbook "${skill}" already exists` }, 409)
  await Bun.write(
    path.join(WORKSPACE_ROOT, ".playbooks", skill, "SKILL.md"),
    // Description is JSON-quoted so a colon or newline cannot inject YAML keys
    // (same posture as skill.ts renderSkillMarkdown).
    `---\nname: ${skill}\ndescription: ${JSON.stringify(description)}\n---\n\n${content}`,
  )
  return c.json({ name: skill, description }, 201)
})

// Revise an imported playbook in WORKSPACE_ROOT/.playbooks. Omitted fields keep
// their current values (the lib backfills from the existing SKILL.md), so a
// description-only edit never wipes the body. Repo-shipped playbooks are read-only.
app.put("/playbooks/:name", async (c) => {
  const input = await c.req.json<{ description?: string; content?: string }>()
  try {
    return c.json(updatePlaybook(c.req.param("name"), input))
  } catch (e: unknown) {
    if (e instanceof PlaybookError) return c.json({ error: e.message }, e.status)
    throw e
  }
})

// Remove an imported playbook. Refused (409) while any matter is still bound to
// it, so a binding never dangles; repo-shipped playbooks are read-only (403).
app.delete("/playbooks/:name", (c) => {
  try {
    deletePlaybook(c.req.param("name"))
    return c.json({ ok: true })
  } catch (e: unknown) {
    if (e instanceof PlaybookError) return c.json({ error: e.message }, e.status)
    throw e
  }
})

app.get("/workflows", (c) => {
  mkdirSync(WORKFLOWS_DIR, { recursive: true })
  return c.json({ dir: WORKFLOWS_DIR, workflows: listWorkflows() })
})

app.post("/workflows", async (c) => {
  const input = await c.req.json<{ label: string; description: string; scope: "matter" | "document"; prompt: string; steps: { agent: string; instructions?: string }[] }>()
  try {
    return c.json(createWorkflow(input), 201)
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.put("/workflows/:name", async (c) => {
  const input = await c.req.json<{ label: string; description: string; scope: "matter" | "document"; prompt: string; steps: { agent: string; instructions?: string }[] }>()
  try {
    return c.json(updateWorkflow(c.req.param("name"), input))
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.delete("/workflows/:name", (c) => {
  try {
    deleteWorkflow(c.req.param("name"))
    return c.json({ ok: true })
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

// The firm's skill library: repo-shipped reference skills (read-only) plus the
// custom ones ingest writes under WORKSPACE_ROOT/.skills. The library directory
// ships alongside the list so the web app can scope its skill-builder chat to it.
app.get("/skills", (c) => {
  mkdirSync(SKILLS_DIR(), { recursive: true })
  return c.json({ dir: SKILLS_DIR(), skills: listSkills() })
})

app.post("/skills", async (c) => {
  const input = await c.req.json<{ name: string; description: string; content: string }>()
  try {
    return c.json(createSkill(input), 201)
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

// Import an uploaded file as a skill: .md keeps its own frontmatter, .docx/.pdf/.txt
// are text-extracted raw for the skill builder to refine afterwards.
app.post("/skills/import", async (c) => {
  const body = await c.req.parseBody()
  const file = body["file"] as File
  try {
    return c.json(await importSkill(file.name, Buffer.from(await file.arrayBuffer())), 201)
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.put("/skills/:name", async (c) => {
  const input = await c.req.json<{ description: string; content: string }>()
  try {
    return c.json(updateSkill(c.req.param("name"), input))
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.delete("/skills/:name", (c) => {
  try {
    deleteSkill(c.req.param("name"))
    return c.json({ ok: true })
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

// Flip a skill on or off (builtin or custom) — the engine permission-denies a
// disabled skill so it leaves every agent's roster without touching the file.
app.patch("/skills/:name", async (c) => {
  const input = await c.req.json<{ enabled: boolean }>()
  try {
    return c.json(setSkillEnabled(c.req.param("name"), input.enabled))
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

// The firm's specialist agents: repo-shipped subagents (read-only) plus the custom
// ones composed by the agent-builder. Custom agents are written to dochaus/agent/
// and registered in dochaus/agents.json; they immediately become available as
// workflow pipeline steps. The library directory ships alongside the list so the
// web app can scope its agent-builder chat to it.
app.get("/agents", (c) => {
  mkdirSync(AGENTS_DIR(), { recursive: true })
  return c.json({ dir: AGENTS_DIR(), agents: listAgents() })
})

app.post("/agents", async (c) => {
  const input = await c.req.json<{ label: string; description: string; instructions: string }>()
  try {
    return c.json(createAgent(input), 201)
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.put("/agents/:name", async (c) => {
  const input = await c.req.json<{ label: string; description: string; instructions: string }>()
  try {
    return c.json(updateAgent(c.req.param("name"), input))
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.delete("/agents/:name", (c) => {
  try {
    deleteAgent(c.req.param("name"))
    return c.json({ ok: true })
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

// Flip an agent on or off (builtin or custom) via agent.<name>.disable in
// dochaus/opencode.json — the same flag the engine's stock agents use.
app.patch("/agents/:name", async (c) => {
  const input = await c.req.json<{ enabled: boolean }>()
  try {
    return c.json(setAgentEnabled(c.req.param("name"), input.enabled))
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.post("/matters", async (c) => {
  const { title, reference, jurisdictions, playbook } = await c.req.json<{
    title: string
    reference?: string
    jurisdictions?: string[]
    playbook?: string
  }>()
  return c.json(createMatter(title, reference, jurisdictions, playbook))
})

app.patch("/matters/:id", async (c) => {
  const { title, reference, jurisdictions, playbook } = await c.req.json<{
    title: string
    reference?: string
    jurisdictions?: string[]
    playbook?: string
  }>()
  return c.json(renameMatter(c.req.param("id"), title, reference, jurisdictions, playbook))
})

app.delete("/matters/:id", (c) => {
  deleteMatter(c.req.param("id"))
  return c.json({ ok: true })
})

app.get("/matters/:id", (c) => {
  const matter = getMatter(c.req.param("id"))
  if (!existsSync(path.join(matter.dir, ".dochaus", "legal.db"))) return c.json({ ...matter, documents: [] })
  const db = openDb(matter.dir)
  const counts = pendingRedlineCounts(db)
  // Surface the pending-redline count per document so the docs rail can badge the
  // documents that have unreviewed changes waiting.
  const documents = (listDocuments(db) as { doc_path: string }[]).map((d) => ({ ...d, pending: counts[d.doc_path] ?? 0 }))
  return c.json({ ...matter, documents })
})

app.post("/matters/:id/documents", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const body = await c.req.parseBody()
  const file = body["file"] as File
  const buffer = Buffer.from(await file.arrayBuffer())
  const result = await ingestDocument(dir, file.name, buffer)
  return c.json(result)
})

// Convert an uploaded .pdf into an editable .docx sibling and index it, so a PDF
// contract can enter the DOCX redline pipeline. The source .pdf is kept; the new
// .docx lands beside it under the same base name. LibreOffice is used when present
// on the host, otherwise a MIT text-only rebuild (see convert.ts).
app.post("/matters/:id/documents/convert", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const name = path.basename(c.req.query("name") ?? "")
  const file = path.join(dir, name)
  if (!name.toLowerCase().endsWith(".pdf") || !existsSync(file)) return c.notFound()
  const docxName = name.replace(/\.pdf$/i, ".docx")
  const result = await ingestDocument(dir, docxName, await pdfToDocx(Buffer.from(await Bun.file(file).bytes())))
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

// The tabular-review grid for a matter: its question-columns and computed cells.
// Rows are the matter's documents, fetched separately, so they are not stored.
app.get("/matters/:id/grid", (c) => c.json(readGrid(matterDir(c.req.param("id")))))

app.put("/matters/:id/grid", async (c) => {
  writeGrid(matterDir(c.req.param("id")), await c.req.json<Grid>())
  return c.json({ ok: true })
})

// Serve a matter's .docx bytes so the web app can render the redline in-browser
// (the viewer converts to HTML client-side via WASM — the file is never uploaded
// anywhere). basename() keeps the lookup inside the matter directory.
app.get("/matters/:id/documents/content", async (c) => {
  const file = path.join(matterDir(c.req.param("id")), path.basename(c.req.query("name") ?? ""))
  if (!existsSync(file)) return c.notFound()
  const mime = file.toLowerCase().endsWith(".pdf")
    ? "application/pdf"
    : "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  return new Response(Bun.file(file).stream(), {
    headers: {
      "Content-Type": mime,
      "Content-Disposition": `inline; filename="${path.basename(file)}"`,
    },
  })
})

// The plain text of a matter document, mammoth-extracted (the same text indexing
// uses). The drafter's read-document tool reads it to convert an existing document
// into a template by replacing every client-specific detail with a placeholder.
// Ships the document's ingest-time injection report (null when clean) so the tool
// can hand the model the text and the warning together.
app.get("/matters/:id/documents/text", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const file = path.join(dir, path.basename(c.req.query("name") ?? ""))
  if (!existsSync(file)) return c.notFound()
  const injection = existsSync(path.join(dir, ".dochaus", "legal.db")) ? getInjectionReport(openDb(dir), file) : null
  return c.json({ text: await extractDocumentText(file, Buffer.from(await Bun.file(file).bytes())), injection })
})

// The global template library. Templates are firm-managed drafting bases shared
// across every matter, stored in WORKSPACE_ROOT/.templates and owned by ingest.
// The library directory ships alongside the list so the web app can scope its
// templates chat to it (the engine sessions run in TEMPLATES_DIR, not a matter).
app.get("/templates", async (c) => c.json({ dir: TEMPLATES_DIR, templates: await listTemplates() }))

app.post("/templates", async (c) => {
  const body = await c.req.parseBody()
  const file = body["file"] as File
  const buffer = Buffer.from(await file.arrayBuffer())
  const name = path.basename(file.name)
  writeFileSync(templatePath(name), buffer)
  const description = typeof body["description"] === "string" ? body["description"] : ""
  setTemplateDescription(name, description)
  const dx = await docxodus()
  const session = dx.openDocxSession(new Uint8Array(buffer), {})
  const placeholders = session.findPlaceholders().map((p) => ({ text: p.match.text, kind: p.kind, hint: p.hint }))
  session.close()
  return c.json({ name, description, placeholders })
})

app.patch("/templates", async (c) => {
  const { description } = await c.req.json<{ description: string }>()
  return c.json(setTemplateDescription(c.req.query("name") ?? "", description))
})

app.delete("/templates", (c) => {
  const name = c.req.query("name") ?? ""
  rmSync(templatePath(name), { force: true })
  removeTemplateDescription(name)
  return c.json({ ok: true })
})

// The plain text of a library template, mammoth-extracted (the same extraction
// the matter document text route uses). The dochaus get-template tool reads it so
// template-builder can revise an existing template instead of recreating it blind.
app.get("/templates/text", async (c) => {
  const file = templatePath(c.req.query("name") ?? "")
  if (!existsSync(file)) return c.notFound()
  return c.json({ text: await extractDocumentText(file, Buffer.from(await Bun.file(file).bytes())) })
})

// Serve a template's .docx bytes so the dochaus draft-document tool can fill it,
// mirroring the matter document content route.
app.get("/templates/content", async (c) => {
  const file = templatePath(c.req.query("name") ?? "")
  if (!existsSync(file)) return c.notFound()
  return new Response(Bun.file(file).stream(), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `inline; filename="${path.basename(file)}"`,
    },
  })
})

// Pending redline proposals for one document — the change list the viewer's
// review panel renders (each as old -> new, with author).
app.get("/matters/:id/redlines", (c) => {
  const dir = matterDir(c.req.param("id"))
  if (!existsSync(path.join(dir, ".dochaus", "legal.db"))) return c.json([])
  const docPath = path.join(dir, path.basename(c.req.query("name") ?? ""))
  return c.json(listPendingRedlines(openDb(dir), docPath))
})

// The redlined view: the clean .docx compared against itself with every pending
// proposal applied, so the viewer renders native tracked changes green/red. Falls
// back to the clean bytes when nothing is pending.
app.get("/matters/:id/documents/redlined", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const file = path.join(dir, path.basename(c.req.query("name") ?? ""))
  if (!existsSync(file)) return c.notFound()
  const rows = existsSync(path.join(dir, ".dochaus", "legal.db"))
    ? listPendingRedlines(openDb(dir), file)
    : []
  const bytes = await buildRedlined(await Bun.file(file).bytes(), rows)
  return new Response(bytes, {
    headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
  })
})

// Accept a single redline: bake it into the canonical .docx (the new accepted
// state, no tracked changes), re-index the document so search reflects the new
// text, and mark it accepted. Re-resolves the anchor text against the live doc, so
// a clause an earlier accept already rewrote fails loudly instead of corrupting.
app.post("/matters/:id/redlines/:rid/accept", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const row = getRedline(openDb(dir), Number(c.req.param("rid")))
  if (!row || row.status !== "pending") return c.json({ error: "Redline not found or already resolved" }, 404)
  const baked = await bake(await Bun.file(row.doc_path).bytes(), [row])
  await ingestDocument(dir, row.doc_name, Buffer.from(baked))
  setRedlineStatus(openDb(dir), row.id, "accepted")
  return c.json({ ok: true })
})

app.post("/matters/:id/redlines/:rid/reject", (c) => {
  const dir = matterDir(c.req.param("id"))
  const row = getRedline(openDb(dir), Number(c.req.param("rid")))
  if (!row || row.status !== "pending") return c.json({ error: "Redline not found or already resolved" }, 404)
  setRedlineStatus(openDb(dir), row.id, "rejected")
  return c.json({ ok: true })
})

// Accept every pending redline on a document in one pass: bake them in document
// order, re-index once, mark all accepted.
app.post("/matters/:id/redlines/accept-all", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const file = path.join(dir, path.basename(c.req.query("name") ?? ""))
  if (!existsSync(file) || !existsSync(path.join(dir, ".dochaus", "legal.db"))) return c.json({ ok: true, accepted: 0 })
  const rows = listPendingRedlines(openDb(dir), file)
  if (rows.length) {
    await ingestDocument(dir, path.basename(file), Buffer.from(await bake(await Bun.file(file).bytes(), rows)))
    for (const row of rows) setRedlineStatus(openDb(dir), row.id, "accepted")
  }
  return c.json({ ok: true, accepted: rows.length })
})

app.post("/matters/:id/redlines/reject-all", (c) => {
  const dir = matterDir(c.req.param("id"))
  const file = path.join(dir, path.basename(c.req.query("name") ?? ""))
  if (!existsSync(path.join(dir, ".dochaus", "legal.db"))) return c.json({ ok: true, rejected: 0 })
  const rows = listPendingRedlines(openDb(dir), file)
  for (const row of rows) setRedlineStatus(openDb(dir), row.id, "rejected")
  return c.json({ ok: true, rejected: rows.length })
})

// Every install gets the repo-shipped starter playbooks, so seed them on boot —
// demo content (the Aldgate Mills matter) stays gated behind `start.sh --demo`,
// which runs seed.ts directly. Dynamic import keeps seed.ts's heavy docx
// dependency out of the request path's module graph.
const { seedPlaybooks } = await import("./seed")
seedPlaybooks()

// Matters indexed before the lexical FTS channel existed (issue #67) get their
// chunks_fts table built and backfilled by openDb's migration here, so the
// search-document tool never opens a legal.db without it. Matters created from
// now on carry the table from first ingest.
for (const matter of listMatters()) {
  if (!existsSync(path.join(matter.dir, ".dochaus", "legal.db"))) continue
  // One corrupt or externally-locked matter DB must not abort the whole
  // service; that matter just stays un-migrated until its next ingest.
  try {
    openDb(matter.dir).close()
  } catch (e) {
    console.error(`failed to migrate ${matter.dir}: ${e instanceof Error ? e.message : e}`)
  }
}

// Bind loopback by default like opencode (packages/opencode/src/cli/network.ts):
// the service is self-hosted alongside the engine and web app, not exposed
// directly. Front it with a reverse proxy to serve beyond localhost.
const port = Number(process.env.INGEST_PORT ?? 4500)
const hostname = process.env.INGEST_HOST ?? "127.0.0.1"
console.log(`doc.haus ingest service listening on http://${hostname}:${port}`)
export default { port, hostname, fetch: app.fetch }
