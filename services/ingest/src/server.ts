import { playbookLibraryRoutes } from "./playbook-library"
import { updateRoutes, managedUpdateOwner } from "./updater"
import { RepoUpdater } from "../../../scripts/updater"
import { documentCopyRoutes } from "./document-copy"
import { templateDraftRoutes } from "./template-drafts"
import { registeredCombinedParts } from "./combined-output"
import { documentLabelRoutes } from "./document-labels"
import { answerDraftRoutes } from "./answer-drafts"
import { readEngineSettings, patchEngineConfig, engineConfigFile } from "./engine-config"
import { setupRoutes } from "./setup"
import { authReuseRoutes } from "./auth-reuse"
import { providerSetupRoutes } from "./provider-setup"
import { Hono } from "hono"
import { attachmentDocumentDisposition, inlineDocumentDisposition } from "./content-disposition"
import { createHash } from "node:crypto"
import { cors } from "hono/cors"
import { HTTPException } from "hono/http-exception"
import {
  openDb,
  listDocuments,
  listPendingRedlines,
  getRedline,
  setRedlineStatus,
  pendingRedlineCounts,
  getExtractionQuality,
  getInjectionReport,
  getDocumentRole,
  type RedlineRow,
} from "./db"
import { readExtractionQuality } from "./extraction-quality"
import {
  ingestDocument,
  extractDocumentText,
  matterDocumentName,
  DeliverableDocumentExistsError,
  SourceDocumentConflictError,
  SourceDocumentExistsError,
} from "./ingest"
import {
  deliverableDocumentName,
  DeliverableConflictError,
  writeDeliverable,
  writeDeliverableRevision,
} from "./deliverable"
import { EmbeddingModelUnavailableError, migrateEmbeddings } from "./embed"
import { migrateSections, migrateStructure } from "./structure"
import { pdfToDocx } from "./convert"
import { buildRedlined, bake, redlineState, redlineStaleReason, RedlinePreviewConflictError } from "./redline"
import { listMatters, createMatter, getMatter, renameMatter, deleteMatter, recordConflicts, matterDir, listJurisdictions, listPlaybooks, updatePlaybook, deletePlaybook, PlaybookError, PLAYBOOK_DIR, MatterValidationError, type Conflicts } from "./matter"
import { listTemplates, templatePath, setTemplateDescription, removeTemplateDescription, TEMPLATES_DIR } from "./template"
import { listWorkflows, createWorkflow, updateWorkflow, deleteWorkflow, NAME_RE, WORKFLOWS_DIR } from "./workflow"
import { listSkills, createSkill, updateSkill, deleteSkill, importSkill, setSkillEnabled, SKILLS_DIR } from "./skill"
import { listAgents, createAgent, updateAgent, deleteAgent, setAgentEnabled, AGENTS_DIR } from "./agent"
import { docxodus } from "./docxodus"
import { listGcpProjects, listAwsProfiles, vertexProbeRoutes } from "./host"
import { parserCapabilitiesRoutes } from "./parser-capabilities-route"
import { checkCellSources, readGrid, recheckPassages, recheckRequest, sourceCheckRequest, writeGrid, type Grid } from "./grid"
import { PreferencesUnavailableError, PreferencesValidationError, readDraftingPreferences, readDraftingTelemetryRetentionPreference, writeDraftingPreferences } from "./preferences"
import { researchConnectionRoutes } from "./research-connections"
import { readLeanPayload, readDetailPayload, readStoragePayload } from "./telemetry"
import { openTelemetryDb, applyRetentionPolicy, clearTelemetry } from "../../../dochaus/lib/telemetry"
import { documentAmendmentRoutes, acceptedWorkingDocxUpdate } from "./document-amendments"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { withParagraphPointers, withStableRevisionDates } from "../../../dochaus/lib/drafting-amendment"
import { DocxTrackedRevisionsError } from "../../../dochaus/lib/docx-amendments"
import { compareVersions } from "../../../dochaus/lib/version-compare"
import { pendingRedlinesForDoc } from "../../../dochaus/lib/redlines"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { ingestEmailDocument, reingestLegacyEmails, removeMatterDocument, withEmailLinks } from "./email-attachments"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { healSourceArtifacts } from "./source-artifacts"
import { MatterSessionDeletionConflictError, MatterSessionDeletionError, deleteMatterSession } from "./matter-session"
import { ChatContinuationConflictError, dispatchChatContinuation, malformedToolContinuationInput } from "./chat-continuation"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { draftingStartProfile, leanDraftingRoutes, leanDraftingStatus, LeanDraftingConflictError, startLeanDraftingRun } from "./lean-drafting"
import {
  DocumentMutationConflictError,
  RegisteredDocumentNotFoundError,
} from "./document-delete"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget, documentRegistryIdentity } from "./document-identity"
import { recoverDocumentMutations } from "./document-mutation"
import { appendAcceptedChangesToNotes } from "./notes-addendum"
import { readLeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
import { replaceableDeliverable } from "../../../dochaus/lib/system-docx"
import { buildCombinedFile } from "../../../dochaus/lib/combined-docx"
import { readLeanState } from "../../../dochaus/workflows/lean/state"
import { documentRedline } from "../../../dochaus/workflows/lean/apply-stage"

const app = new Hono()

// A model that cannot load fails only the request that needed it, with a
// message the web can show; every other error keeps Hono's default handling.
app.onError((error, c) => {
  if (error instanceof PreferencesValidationError) return c.json({ error: error.message }, 400)
  if (error instanceof PreferencesUnavailableError) return c.json({ error: error.message }, 503)
  if (error instanceof MatterValidationError) return c.json({ error: error.message }, 400)
  if (error instanceof EmbeddingModelUnavailableError)
    return c.json({ error: error.message, code: "embedding_model_unavailable" }, 503)
  if (error instanceof HTTPException) return error.getResponse()
  console.error(error)
  return c.text("Internal Server Error", 500)
})

// transformers.js leaves some model-download promises unawaited when a sibling
// download fails first (for example a reset connection mid-download); their
// later rejection would otherwise terminate the whole service. The request that
// triggered the load has already failed through loadEmbeddingModel, so record
// the stray rejection and keep serving.
process.on("unhandledRejection", (reason) => {
  console.error("ingest: unhandled rejection; service continues", {
    ...diagnosticErrorMetadata(reason),
    ...(reason instanceof Error && "code" in reason && typeof reason.code === "string" ? { code: reason.code } : {}),
  })
})

// Match the legal engine's CORS posture: allow
// only localhost/127.0.0.1 origins on any port, not a wildcard. The web app runs
// on localhost; a wildcard would let any site in a user's browser hit this
// service. Non-browser callers (curl, the opencode server) send no Origin and are
// unaffected.
app.use(
  "*",
  cors({ origin: (origin) => (/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin) ? origin : null) }),
)

const updateRoot = path.resolve(import.meta.dir, "../../..")
const updateRuntime = process.env.DOCHAUS_RUNTIME_DIR ?? path.join(updateRoot, ".dochaus-runtime")
app.route("/", updateRoutes(new RepoUpdater(updateRoot, updateRuntime), async () => {
  if (listMatters().some(matter => leanDraftingStatus(matter.dir)?.in_flight)) return true
  const engine = process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
  const response = await fetch(`${engine}/api/session/active`, { signal: AbortSignal.timeout(5000) })
  if (!response.ok) throw new Error("Could not check ongoing work. Please try again when the app is connected.")
  const active = await response.json()
  if (typeof active !== "object" || active === null || !("data" in active) || !active.data || typeof active.data !== "object") throw new Error("Could not check ongoing work. Please try again.")
  return Object.keys(active.data).length > 0
}, () => managedUpdateOwner(updateRoot, updateRuntime)))

app.route("/", documentLabelRoutes(matterDir))
app.route("/", answerDraftRoutes(matterDir))

app.route("/", authReuseRoutes())
app.route("/", providerSetupRoutes())

app.route("/", setupRoutes({ directory: path.dirname(engineConfigFile()) }))

app.get("/engine/config", c => c.json(readEngineSettings()))
app.patch("/engine/config", async c => {
  try {
    return c.json(patchEngineConfig(await c.req.json()))
  } catch (error) {
    return c.json({ error: error instanceof Error ? error.message : "Invalid engine settings." }, 400)
  }
})

// What the host's cloud sign-ins can see, for the provider-setup comboboxes in
// the web settings (GCP projects via gcloud ADC, AWS profiles via ~/.aws).
// Empty lists when the host has no such sign-in (e.g. gcloud not installed).
app.get("/host/gcp-projects", async (c) => c.json({ projects: await listGcpProjects().catch(() => []) }))

app.get("/host/aws-profiles", async (c) => c.json({ profiles: await listAwsProfiles() }))

app.route("/host", parserCapabilitiesRoutes())

app.route("/host", vertexProbeRoutes())

app.get("/matters", (c) => c.json(listMatters()))

// Firm-wide drafting preferences (Settings → Drafting). Saving re-renders the
// standing-instructions markdown the engine reads on every turn (see
// preferences.ts), so a change applies from the assistant's next reply.
app.get("/preferences", (c) => c.json(readDraftingPreferences()))

app.put("/preferences", async (c) => {
  const input = await c.req.json().catch(() => undefined)
  return c.json(writeDraftingPreferences(input))
})
app.route("/research", researchConnectionRoutes)

// The jurisdiction packs a matter can be assigned, read from the dochaus config
// layer. The web app populates its matter-creation dropdown from this (issue #18).
app.get("/jurisdictions", (c) => c.json(listJurisdictions()))

// The playbook skills a matter can be reviewed against: the repo-shipped ones plus
// any the firm imported into WORKSPACE_ROOT/.playbooks. The web app offers these as
// the matter's playbook choice; create-playbook reads it to dedupe before writing.
app.get("/playbooks", (c) => c.json(listPlaybooks()))
app.route("/", playbookLibraryRoutes())

// Import a firm playbook skill into WORKSPACE_ROOT/.playbooks. Ingest is the single
// writer over WORKSPACE_ROOT (same rationale as the .templates library), so the
// dochaus create-playbook tool POSTs here rather than writing the file itself. The
// frontmatter (name + quoted description) is added here; the body is supplied raw.
app.post("/playbooks", async (c) => {
  const input = await c.req.json<{ name: string; description: string; content: string }>()
  if (!input || typeof input !== "object") return c.json({ error: "Supply a playbook name, description and non-empty content" }, 400)
  const { name, description, content } = input
  if (typeof name !== "string" || typeof description !== "string" || typeof content !== "string" || !content.trim())
    return c.json({ error: "Supply a playbook name, description and non-empty content" }, 400)
  if (!name.startsWith("playbook-")) return c.json({ error: 'Playbook name must start with "playbook-"' }, 400)
  const skill = path.basename(name)
  // Same slug rule skill.ts enforces — the name goes onto an unquoted YAML
  // `name:` line below, so only lowercase/hyphen slugs are safe to write.
  if (!NAME_RE.test(skill) || !/^playbook-[a-z0-9]/.test(skill)) return c.json({ error: `"${skill}" is not a valid playbook name (lowercase, hyphenated)` }, 400)
  if (listPlaybooks().some((p) => p.name === skill)) return c.json({ error: `playbook "${skill}" already exists` }, 409)
  await Bun.write(
    path.join(PLAYBOOK_DIR(), skill, "SKILL.md"),
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
  if (!input || typeof input !== "object") return c.json({ error: "Supply the playbook fields to update" }, 400)
  if ((input.description !== undefined && typeof input.description !== "string") ||
      (input.content !== undefined && typeof input.content !== "string"))
    return c.json({ error: "Playbook description and content must be text" }, 400)
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
  const { title, reference, jurisdictions, playbook, client, posture } = await c.req.json<{
    title: string
    reference?: string
    jurisdictions?: string[]
    playbook?: string
    client?: string
    posture?: string
  }>()
  return c.json(createMatter(title, reference, jurisdictions, playbook, client, posture))
})

// A body without a client or posture key keeps the stored one; "" clears it.
app.patch("/matters/:id", async (c) => {
  const body = await c.req.json<{
    title: string
    reference?: string
    jurisdictions?: string[]
    playbook?: string
    client?: string
    posture?: string
  }>()
  const { title, reference, jurisdictions, playbook } = body
  return c.json(renameMatter(c.req.param("id"), title, reference, jurisdictions, playbook,
    "client" in body ? body.client : undefined, "posture" in body ? body.posture : undefined))
})

// The conflicts intake gate: records that a conflicts check happened and how it
// exited (cleared / pending / bypassed) — it never judges conflicts itself. The
// dochaus plugin keys reviewer-note framing and edit-tool blocking off this
// record deterministically.
app.post("/matters/:id/conflicts", async (c) => {
  const input = await c.req.json<{ status: Conflicts["status"]; by?: string; owner?: string; due?: string; reason?: string }>()
  try {
    return c.json(recordConflicts(c.req.param("id"), input))
  } catch (e: any) {
    return c.json({ error: e.message }, e.status ?? 400)
  }
})

app.delete("/matters/:id", (c) => {
  deleteMatter(c.req.param("id"))
  return c.json({ ok: true })
})

app.get("/matters/:id", async (c) => {
  const matter = getMatter(c.req.param("id"))
  return withArtifactMutationLocksAsync(matter.dir, [".dochaus/source-corpus"], async () => {
    if (!existsSync(path.join(matter.dir, ".dochaus", "legal.db"))) return c.json({ ...matter, documents: [] })
    const db = openDb(matter.dir)
    try {
      const counts = pendingRedlineCounts(db)
      // Surface the pending-redline count per document so the docs rail can badge the
      // documents that have unreviewed changes waiting.
      // Each source's extraction quality (#239) goes out as `quality`; one
      // ingested before the gate has none.
      const documents = withEmailLinks(db, (listDocuments(db) as { doc_path: string; name: string; extraction_quality: string | null }[]).map(({ extraction_quality, ...d }) => ({
        ...d,
        ...(extraction_quality ? { quality: readExtractionQuality(extraction_quality) } : {}),
        pending: counts[documentRegistryIdentity(d.doc_path)] ?? 0,
      })))
      return c.json({ ...matter, documents })
    } finally {
      db.close()
    }
  })
})

// Destructive conversation operations stay behind the matter-aware web
// service; see deleteMatterSession for the drafting-owner guard.
app.delete("/matters/:id/sessions/:sessionID", async (c) => {
  try {
    return c.json(await deleteMatterSession(
      matterDir(c.req.param("id")),
      c.req.param("sessionID"),
    ))
  } catch (error) {
    if (error instanceof MatterSessionDeletionConflictError)
      return c.json({ error: error.message }, 409)
    if (error instanceof MatterSessionDeletionError)
      return c.json({ error: error.message }, 502)
    throw error
  }
})

app.get("/matters/:id/telemetry", (c) =>
  c.json(readLeanPayload(matterDir(c.req.param("id")), {
    session: c.req.query("session") ?? undefined,
    rateOverrides: readDraftingPreferences().telemetry.rate_overrides,
  })),
)

app.get("/matters/:id/telemetry/detail", (c) => c.json(readDetailPayload(matterDir(c.req.param("id")), readDraftingPreferences().telemetry.rate_overrides)))

app.get("/matters/:id/telemetry/storage", (c) => c.json(readStoragePayload([matterDir(c.req.param("id"))], readDraftingPreferences().telemetry.rate_overrides)))

app.get("/matters/:id/drafting-status", (c) => {
  return c.json(leanDraftingStatus(matterDir(c.req.param("id"))) ?? null)
})

app.post("/matters/:id/drafting-runs/start", async (c) => {
  const input: {
    request?: string
    agent?: string
    model?: { providerID?: string; modelID?: string }
    profile?: unknown
    // An upgraded quick draft: an unverified working draft, never request text.
    workingDraft?: unknown
  } = await c.req.json().catch(() => ({}))
  const profile = draftingStartProfile(input.profile)
  if (!profile) return c.json({ error: "Drafting profile must be \"lean\".", code: "invalid_drafting_profile" }, 400)
  if (
    typeof input.request !== "string" ||
    !input.request.trim() ||
    input.request.replace(/\r\n/g, "\n").trim().length > 20_000 ||
    input.agent !== "draft" ||
    !input.model ||
    typeof input.model.providerID !== "string" ||
    !input.model.providerID.trim() ||
    typeof input.model.modelID !== "string" ||
    !input.model.modelID.trim()
  )
    return c.json({ error: "Drafting start requires a resolved assistant and request." }, 400)
  if (input.workingDraft !== undefined && typeof input.workingDraft !== "string")
    return c.json({ error: "A working draft must be text." }, 400)
  const model = {
    providerID: input.model.providerID,
    modelID: input.model.modelID,
  }
  try {
    return c.json(await startLeanDraftingRun(matterDir(c.req.param("id")), { request: input.request, agent: input.agent, model, workingDraft: input.workingDraft }))
  } catch (error) {
    if (error instanceof LeanDraftingConflictError) return c.json({ error: error.message, code: error.code }, 409)
    throw error
  }
})

app.post("/matters/:id/continuations/dispatch", async (c) => {
  const input = await c.req.json<{
    kind?: unknown
    sessionID?: unknown
    progressToken?: unknown
  }>().catch(() => undefined)
  if (!input) return c.json({ error: "The continuation dispatch request is invalid." }, 400)
  try {
    if (input.kind !== "malformed-tool")
      return c.json({ error: "The continuation dispatch kind is invalid." }, 400)
    if (typeof input.sessionID !== "string" || typeof input.progressToken !== "string")
      return c.json({ error: "The interrupted-turn continuation identity is invalid." }, 400)
    const failedMessageID = input.progressToken.match(/^malformed:([A-Za-z0-9_-]{1,512})$/)?.[1]
    if (!failedMessageID)
      return c.json({ error: "The interrupted-turn continuation identity is invalid." }, 400)
    return c.json(await dispatchChatContinuation(
      matterDir(c.req.param("id")),
      malformedToolContinuationInput(input.sessionID, failedMessageID),
      { requireIdleOwner: true },
    ))
  } catch (error) {
    if (error instanceof ChatContinuationConflictError) return c.json({ error: error.message }, 409)
    throw error
  }
})

app.route("/matters", documentAmendmentRoutes({ directory: matterDir }))
app.route("/matters", leanDraftingRoutes({ directory: matterDir }))

app.get("/telemetry/storage", (c) => c.json(readStoragePayload(listMatters().map((m) => m.dir), readDraftingPreferences().telemetry.rate_overrides)))

app.delete("/matters/:id/telemetry", (c) => {
  const dir = matterDir(c.req.param("id"))
  if (!existsSync(path.join(dir, ".dochaus", "telemetry.db"))) return c.json({ ok: true })
  const db = openTelemetryDb(dir)
  try {
    clearTelemetry(db)
  } finally {
    db.close()
  }
  return c.json({ ok: true })
})

app.post("/matters/:id/documents", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const body = await c.req.parseBody()
  const file = body["file"]
  if (!(file instanceof File)) return c.json({ error: "A source document file is required." }, 400)
  const name = matterDocumentName(file.name)
  if (!name) return c.json({ error: "The source document filename or extension is not supported." }, 400)
  // A same-name source is never replaced silently: without an `existing` choice
  // the upload is refused with source_exists so the lawyer can choose.
  const existing = body["existing"]
  if (existing !== undefined && existing !== "replace" && existing !== "keep-both")
    return c.json({ error: "Choose replace or keep-both for an existing document." }, 400)
  const buffer = Buffer.from(await file.arrayBuffer())
  try {
    // An email brings its attachments in after it (email-attachments.ts). A
    // lawyer's upload is never an attachment, even under an attachment's name.
    if (name.toLowerCase().endsWith(".eml")) return c.json(await ingestEmailDocument(dir, name, buffer, { existing }))
    return c.json(await ingestDocument(dir, name, buffer, { existing, links: { parentName: null } }))
  } catch (error) {
    if (error instanceof SourceDocumentExistsError)
      return c.json({ error: error.message, code: "source_exists", name: error.existing }, 409)
    if (error instanceof DeliverableDocumentExistsError)
      return c.json({ error: error.message, code: "deliverable_exists", name: error.existing }, 409)
    if (error instanceof SourceDocumentConflictError) return c.json({ error: error.message }, 409)
    throw error
  }
})

// Generated documents are matter deliverables, not evidence for the matter.
// Register and write them so the Documents rail/content route can serve them,
// without extraction, chunking, embedding, or rebuilding source artifacts.
app.post("/matters/:id/deliverables", async (c) => {
  const body = await c.req.parseBody()
  const file = body["file"]
  if (!(file instanceof File)) return c.json({ error: "A deliverable file is required." }, 400)
  if (!deliverableDocumentName(file.name))
    return c.json({ error: "The deliverable filename or extension is not supported." }, 400)
  const preconditionKind = c.req.header("x-dochaus-deliverable-precondition")
  const priorSha256 = c.req.header("x-dochaus-deliverable-prior-sha256")
  if (
    (preconditionKind &&
      preconditionKind !== "create" &&
      preconditionKind !== "replace") ||
    (preconditionKind === "create" && priorSha256) ||
    (preconditionKind === "replace" &&
      !/^[a-f0-9]{64}$/.test(priorSha256 ?? "")) ||
    (!preconditionKind && priorSha256)
  )
    return c.json({ error: "The deliverable write precondition is invalid." }, 400)
  const precondition = preconditionKind === "create"
    ? { kind: "create" as const }
    : preconditionKind === "replace"
      ? { kind: "replace" as const, prior_sha256: priorSha256! }
      : undefined
  try {
    return c.json(await writeDeliverable(
      matterDir(c.req.param("id")),
      file.name,
      Buffer.from(await file.arrayBuffer()),
      { precondition },
    ))
  } catch (e) {
    if (e instanceof DeliverableConflictError) return c.json({ error: e.message }, 409)
    throw e
  }
})

app.route("/", documentCopyRoutes(matterDir))
app.route("/", templateDraftRoutes(matterDir))

// Persist an edit to an already registered document while preserving its role.
// Editing a generated deliverable must never promote it into source retrieval.
app.post("/matters/:id/documents/revisions", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const body = await c.req.parseBody()
  const file = body["file"]
  if (!(file instanceof File)) return c.json({ error: "A revised document file is required." }, 400)
  const sourceName = matterDocumentName(file.name)
  const deliverableName = deliverableDocumentName(file.name)
  if (!sourceName && !deliverableName)
    return c.json({ error: "The revised document filename or extension is not supported." }, 400)
  let target: ReturnType<typeof resolveMatterDocumentTarget>
  try {
    target = resolveMatterDocumentTarget(dir, sourceName ?? deliverableName!)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) return c.json({ error: error.message }, 409)
    throw error
  }
  const db = openDb(dir)
  const role = getDocumentRole(db, target.docPath)
  db.close()
  if (!role) return c.json({ error: "The document is not registered in this matter." }, 404)
  if ((role === "source" && !sourceName) || (role === "deliverable" && !deliverableName))
    return c.json({ error: `The filename extension is not valid for a ${role} document.` }, 400)
  try {
    const priorSha256 = c.req.header("x-document-prior-sha256")
    const bytes = Buffer.from(await file.arrayBuffer())
    // Under the document's lock, which an accept holds from building its run's
    // delivered redline to writing it, so a revision never lands in between. The
    // revision writer takes the same lock; taking it here also keeps the
    // precondition below and the working-revision update under it. A revision
    // of a part of the run's saved combined file rebuilds that file as an accept
    // does, under the same extra locks (runCombinedLock).
    return await withArtifactMutationLocksAsync(dir, [target.name, ...runCombinedLock(dir, target.docPath), ".dochaus/source-corpus"], async () => {
      if (readWorkingDocx(dir,target.docPath) && !/^[a-f0-9]{64}$/.test(priorSha256 ?? ""))
        return c.json({error:"Provide the exact current Word hash in x-document-prior-sha256 before replacing an accepted revision."},409)
      // A deliverable's prior hash is checked by its writer; a source's here.
      if (priorSha256 && role === "source" && sha256Hex(readFileSync(target.docPath)) !== priorSha256)
        return c.json({error:"The document changed after it was read. Read it again before editing."},409)
      const combined = await prepareCombinedFile(dir, target.docPath)
      return c.json({ ...await writeDocumentRevision(dir, target.name, bytes,[],priorSha256), ...await rebuildCombinedFile(combined) })
    })
  } catch (error) {
    if (error instanceof SourceDocumentConflictError || error instanceof DeliverableConflictError)
      return c.json({ error: error.message }, 409)
    // An accepted working revision must be readable paragraph by paragraph.
    if (error instanceof DocxTrackedRevisionsError)
      return c.json({ error: "This Word file contains tracked changes. Accept or reject them in Word, then upload it again." }, 409)
    throw error
  }
})

// Convert an uploaded .pdf into an editable .docx sibling and index it, so a PDF
// contract can enter the DOCX redline pipeline. The source .pdf is kept; the new
// .docx lands beside it under the same base name, or the next free one when that
// name is taken, so a conversion never overwrites a document. LibreOffice is used
// when present on the host, otherwise a MIT text-only rebuild (see convert.ts).
app.post("/matters/:id/documents/convert", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const name = matterDocumentName(c.req.query("name") ?? "")
  if (!name?.toLowerCase().endsWith(".pdf")) return c.notFound()
  const requestedDocxName = name.replace(/\.pdf$/i, ".docx")
  return withArtifactMutationLocksAsync(dir, [name, requestedDocxName, ".dochaus/source-corpus"], async () => {
    let target: ReturnType<typeof resolveMatterDocumentTarget>
    try {
      target = resolveMatterDocumentTarget(dir, name)
    } catch (error) {
      if (error instanceof MatterDocumentIdentityError) return c.json({ error: error.message }, 409)
      throw error
    }
    const file = target.docPath
    if (!existsSync(file)) return c.notFound()
    const db = openDb(dir)
    const role = getDocumentRole(db, file)
    db.close()
    if (role !== "source") return c.json({ error: `${target.name} is not a registered source document.` }, 409)
    const docxName = target.name.replace(/\.pdf$/i, ".docx")
    const result = await ingestDocument(dir, docxName, await pdfToDocx(Buffer.from(await Bun.file(file).bytes())), { existing: "keep-both" })
    return c.json(result)
  })
})

// Remove a document: delete the source .docx and its index rows. basename()
// keeps the lookup inside the matter directory, matching the content route.
app.delete("/matters/:id/documents", async (c) => {
  const requested = c.req.query("name") ?? ""
  if (!matterDocumentName(requested) && !deliverableDocumentName(requested))
    return c.json({ error: "The document filename or extension is not supported." }, 400)
  try {
    // Removing an email also removes its attachments unless attachments=keep.
    return c.json(await removeMatterDocument(matterDir(c.req.param("id")), requested, c.req.query("attachments") === "keep"))
  } catch (error) {
    if (error instanceof RegisteredDocumentNotFoundError) return c.json({ error: error.message }, 404)
    if (error instanceof DocumentMutationConflictError) return c.json({ error: error.message }, 409)
    throw error
  }
})

// The tabular-review grid for a matter: its question-columns and computed cells.
// Rows are the matter's documents, fetched separately, so they are not stored.
app.get("/matters/:id/grid", (c) => c.json(readGrid(matterDir(c.req.param("id")))))

app.put("/matters/:id/grid", async (c) => {
  writeGrid(matterDir(c.req.param("id")), await c.req.json<Grid>())
  return c.json({ ok: true })
})

// One deterministic re-check pass for a negative grid cell, scoped to its own
// document: read-only candidates and the scope searched (see grid.ts).
app.post("/matters/:id/grid/recheck", async (c) => {
  const request = recheckRequest(await c.req.json().catch(() => undefined))
  if (!request) return c.json({ error: "A re-check needs one document name, the question and the passages already read." }, 400)
  return c.json(recheckPassages(matterDir(c.req.param("id")), request))
})

// Whether each cited cell's excerpt is still the document's text at its
// offsets (see checkCellSources). Grid citations come from the matter's indexed
// sources, so only a file registered as a source resolves; a deliverable is
// never indexed or cited, and an unregistered file reads as missing.
app.post("/matters/:id/grid/sources", async (c) => {
  const request = sourceCheckRequest(await c.req.json().catch(() => undefined))
  if (!request) return c.json({ error: "A source check needs a list of cited cells." }, 400)
  const dir = matterDir(c.req.param("id"))
  const db = existsSync(path.join(dir, ".dochaus", "legal.db")) ? openDb(dir) : undefined
  const targets = (() => {
    try {
      return new Map(request.cells.map((cell) => {
        const name = matterDocumentName(cell.documentName)
        const target = name ? readDocumentTarget(dir, name) : undefined
        const source = (() => {
          try {
            return !!db && !!target && existsSync(target.docPath) && getDocumentRole(db, target.docPath) === "source"
          } catch {
            return false
          }
        })()
        return [cell.documentName, source ? target : undefined] as const
      }))
    } finally {
      db?.close()
    }
  })()
  const names = [...new Set([...targets.values()].flatMap((target) => (target ? [target.name] : [])))]
  return withArtifactMutationLocksAsync(dir, names, async () => c.json(await checkCellSources(request, (name) => targets.get(name))))
})

// Serve a matter's .docx bytes so the web app can render the redline in-browser
// (the viewer converts to HTML client-side via WASM — the file is never uploaded
// anywhere). basename() keeps the lookup inside the matter directory.
app.get("/matters/:id/documents/content", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const name = registeredDocumentName(c.req.query("name") ?? "")
  if (!name) return c.notFound()
  return withArtifactMutationLocksAsync(dir, [name], async () => {
    const target = readDocumentTarget(dir, name)
    if (!target || !existsSync(target.docPath)) return c.notFound()
    const mime = target.docPath.toLowerCase().endsWith(".pdf")
      ? "application/pdf"
      : "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return new Response(await Bun.file(target.docPath).bytes(), {
      headers: {
        "Content-Type": mime,
        "Content-Disposition": inlineDocumentDisposition(target.name),
      },
    })
  })
})

// The plain text of a matter document, mammoth-extracted (the same text indexing
// uses). The drafter's read-document tool reads it to convert an existing document
// into a template by replacing every client-specific detail with a placeholder.
// Ships the document's ingest-time injection report (null when clean) and its
// extraction quality so the tool can hand the model the text and the warnings
// together.
app.get("/matters/:id/documents/text", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const name = registeredDocumentName(c.req.query("name") ?? "")
  if (!name) return c.notFound()
  return withArtifactMutationLocksAsync(dir, [name], async () => {
    const target = readDocumentTarget(dir, name)
    if (!target || !existsSync(target.docPath)) return c.notFound()
    const db = existsSync(path.join(dir, ".dochaus", "legal.db")) ? openDb(dir) : undefined
    const { injection, quality } = (() => {
      if (!db) return { injection: null, quality: undefined }
      try {
        return { injection: getInjectionReport(db, target.docPath), quality: getExtractionQuality(db, target.docPath) }
      } finally {
        db.close()
      }
    })()
    return c.json({
      text: await extractDocumentText(target.name, Buffer.from(await Bun.file(target.docPath).bytes())),
      injection,
      // The extraction quality gate's result (#239); absent for a document
      // ingested before it.
      ...(quality ? { quality } : {}),
    })
  })
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
  mkdirSync(TEMPLATES_DIR, { recursive: true })
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
      "Content-Disposition": inlineDocumentDisposition(path.basename(file)),
    },
  })
})

// Pending redline proposals for one document — the change list the viewer's
// review panel renders (each as old -> new, with author).
app.get("/matters/:id/redlines", async (c) => {
  const dir = matterDir(c.req.param("id"))
  if (!existsSync(path.join(dir, ".dochaus", "legal.db"))) return c.json([])
  const name = registeredDocumentName(c.req.query("name") ?? "")
  if (!name) return c.json([])
  return withArtifactMutationLocksAsync(dir, [name], async () => {
    const target = readDocumentTarget(dir, name)
    if (!target || !existsSync(target.docPath)) return c.json([])
    const db = openDb(dir)
    try {
      const rows = listPendingRedlines(db, target.docPath)
      const state = await redlineState(await Bun.file(target.docPath).bytes(),rows)
      const protectedWord = Boolean(readWorkingDocx(dir,target.docPath))
      return c.json(rows.map((row) => {
        const stale_reason = redlineStaleReason(row,state,protectedWord)
        return {...row,...(stale_reason ? {stale:true,stale_reason} : {})}
      }))
    } finally {
      db.close()
    }
  })
})

// The redlined view: the clean .docx compared against itself with every pending
// proposal applied, so the viewer renders native tracked changes green/red. Falls
// back to the clean bytes when nothing is pending. Stale proposals (the ones the
// change list marks stale and refuses to accept) are left out: they cannot be
// replayed against this revision, and they stay pending and visible as stale
// until the reviewer rejects them. Nothing here changes a proposal's status.
app.get("/matters/:id/documents/redlined", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const name = registeredDocumentName(c.req.query("name") ?? "")
  if (!name) return c.notFound()
  return withArtifactMutationLocksAsync(dir, [name], async () => {
    const target = readDocumentTarget(dir, name)
    if (!target || !existsSync(target.docPath)) return c.notFound()
    const original = await Bun.file(target.docPath).bytes()
    const protectedWord = Boolean(readWorkingDocx(dir, target.docPath))
    const db = existsSync(path.join(dir, ".dochaus", "legal.db")) ? openDb(dir) : undefined
    const pending = (() => {
      if (!db) return []
      try {
        return listPendingRedlines(db, target.docPath)
      } finally {
        db.close()
      }
    })()
    const state = await redlineState(original, pending)
    const rows = pending.filter((row) => !redlineStaleReason(row, state, protectedWord))
    try {
      const bytes = await buildRedlined(original, rows, state)
      return new Response(new Uint8Array(bytes), {
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      })
    } catch(error) {
      if (error instanceof RedlinePreviewConflictError) return c.json({error:error.message},409)
      throw error
    }
  })
})

// Compare two versions of a matter document, left (the earlier) with right (the
// later): a Word Compare-style blackline (format=docx, the default) or its
// paragraph changes, the edits right made without tracked changes and what
// became of each of left's tracked proposals (format=json). The stored bytes
// are compared; pending doc.haus proposals are counted, not applied. Read
// only: no model call and nothing written to the matter, so no artifact lock
// (it would create its lock folder): document writes replace
// the whole file by rename, so each read sees one complete version, and legal.db
// is opened read-only, never created or migrated.
app.get("/matters/:id/compare", async (c) => {
  const refuse = (status: 400 | 404 | 422, error: string, message: string) => c.json({ error, message }, status)
  const unknown = (name: string) => refuse(404, "not_found", `${name} is not a document in this matter.`)
  const dir = matterDir(c.req.param("id"))
  const requested = [c.req.query("left")?.trim() ?? "", c.req.query("right")?.trim() ?? ""]
  if (!requested[0] || !requested[1]) return refuse(400, "missing_name", "Choose two documents to compare.")
  if (!existsSync(dir)) return refuse(404, "not_found", "This matter does not exist.")
  const [leftName, rightName] = requested.map((name) => registeredDocumentName(name))
  if (!leftName) return unknown(requested[0])
  if (!rightName) return unknown(requested[1])
  // A symbolic-link alias, inside or out of the matter, resolves to no target.
  const left = readDocumentTarget(dir, leftName)
  if (!left || !existsSync(left.docPath)) return unknown(leftName)
  const right = readDocumentTarget(dir, rightName)
  if (!right || !existsSync(right.docPath)) return unknown(rightName)
  for (const target of [left, right])
    if (path.extname(target.name).toLowerCase() !== ".docx")
      return refuse(400, "not_docx", `${target.name} is not a Word document. Use Convert to DOCX on it first, then compare the converted file.`)
  if (left.identity === right.identity) return refuse(400, "same_document", `Both sides are ${left.name}. Choose two different documents.`)
  const stem = (name: string) => path.parse(name).name
  // A label becomes a Word revision author: one line, capped as the compare tool caps it.
  const oneLine = (text: string) => Array.from(text.replace(/[\p{Cc}\s]+/gu, " ").trim()).slice(0, 120).join("").trim()
  const label = (key: string, name: string) => oneLine(c.req.query(key) ?? "") || oneLine(stem(name))
  const flag = (key: string) => c.req.query(key) !== "0"
  const options = { ignoreFormatting: flag("ignoreFormatting"), detectMoves: flag("moves"), headersFooters: flag("headers") }
  const labels = { left: label("leftLabel", left.name), right: label("rightLabel", right.name) }
  // A document deleted, or between the two renames of a replacement, since it was resolved.
  const read = (file: string) => Bun.file(file).bytes().catch((error: unknown) => {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") return undefined
    throw error
  })
  const leftBytes = await read(left.docPath)
  if (!leftBytes) return unknown(left.name)
  const rightBytes = await read(right.docPath)
  if (!rightBytes) return unknown(right.name)
  const pending = { left: pendingRedlinesForDoc(dir, left.docPath).length, right: pendingRedlinesForDoc(dir, right.docPath).length }
  const comparison = await compareVersions(
    { bytes: leftBytes, label: labels.left },
    { bytes: rightBytes, label: labels.right },
    options,
    new Date().toISOString(),
  ).catch((error: unknown) => {
    console.warn("[compare] comparison failed", diagnosticErrorMetadata(error))
    return undefined
  })
  if (!comparison)
    return refuse(422, "compare_failed", `${left.name} and ${right.name} could not be compared. One of them may be damaged or not a standard Word file: open it in Word, save it again and retry.`)
  const blacklineName = `Blackline - ${stem(left.name)} vs ${stem(right.name)}.docx`
  if (c.req.query("format") === "json")
    return c.json({
      left: { name: left.name, label: labels.left, trackedChanges: comparison.left, pendingProposals: pending.left },
      right: { name: right.name, label: labels.right, trackedChanges: comparison.right, pendingProposals: pending.right },
      options,
      attribution: comparison.attribution,
      changes: comparison.changes,
      untracked: comparison.untracked,
      proposals: comparison.proposals,
      blacklineName,
    })
  return new Response(new Uint8Array(comparison.blackline), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": c.req.query("download") === "1" ? attachmentDocumentDisposition(blacklineName) : inlineDocumentDisposition(blacklineName),
    },
  })
})

// Accept a single redline: bake it into the canonical .docx (the new accepted
// state, no tracked changes) and mark it accepted. Source revisions are
// re-indexed; deliverable revisions remain outside source retrieval. Re-resolves
// the anchor text against the live doc, so a clause an earlier accept already
// rewrote fails loudly instead of corrupting.
app.post("/matters/:id/redlines/:rid/accept", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const initialDb = openDb(dir)
  const row = getRedline(initialDb, Number(c.req.param("rid")))
  initialDb.close()
  if (!row || row.status !== "pending") return c.json({ error: "Redline not found or already resolved" }, 404)
  return withArtifactMutationLocksAsync(dir, [row.doc_name, ...runRedlineLock(dir, row.doc_path), ...runNotesLock(dir), ...runCombinedLock(dir, row.doc_path), ".dochaus/source-corpus"], async () => {
    const db = openDb(dir)
    const current = getRedline(db, row.id)
    db.close()
    if (!current || current.status !== "pending")
      return c.json({ error: "Redline not found or already resolved" }, 404)
    const combined = await prepareCombinedFile(dir, current.doc_path)
    let redline: PreparedRedline
    try {
      const original = await Bun.file(current.doc_path).bytes()
      const prior = createHash("sha256").update(original).digest("hex")
      const working = readWorkingDocx(dir,current.doc_path)
      const state = await redlineState(original,[current])
      const stale = redlineStaleReason(current,state,Boolean(working))
      if (stale || working && working.source_sha256 !== prior) return c.json({error:stale ?? "The working Word revision is inconsistent. Refresh it through the revision API."},409)
      const baked = documentRevisionBytes(dir,current.doc_name,Buffer.from(await bake(original,[current],state))).bytes
      redline = await prepareDeliveredRedline(dir,current.doc_path,baked)
      await writeDocumentRevision(dir,current.doc_name,baked,[current.id],prior,current.target_json != null)
    } catch (error) {
      console.warn("[redline] acceptance refused",diagnosticErrorMetadata(error))
      return c.json({error:"Acceptance could not be confirmed. Refresh the document and change list before trying again."},409)
    }
    // In order: the redline, the notes addendum, then the combined file built from the parts as written.
    return c.json({ ok: true, ...await writeDeliveredRedline(redline), ...await recordAcceptedChangesInRunNotes(dir, row.doc_path, [current]), ...await rebuildCombinedFile(combined) })
  })
})

app.post("/matters/:id/redlines/:rid/reject", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const initialDb = openDb(dir)
  const row = getRedline(initialDb, Number(c.req.param("rid")))
  initialDb.close()
  if (!row || row.status !== "pending") return c.json({ error: "Redline not found or already resolved" }, 404)
  return withArtifactMutationLocksAsync(dir, [row.doc_name, ...rejectCombinedLock(dir, row.doc_path)], async () => {
    const db = openDb(dir)
    let current: RedlineRow | null
    try {
      current = getRedline(db, row.id)
    } finally {
      db.close()
    }
    if (!current || current.status !== "pending")
      return c.json({ error: "Redline not found or already resolved" }, 404)
    // As an accept does: a rebuild of the saved combined file skipped earlier is made now.
    const combined = await prepareCombinedFile(dir, current.doc_path, { rejected: true })
    const update = openDb(dir)
    try {
      setRedlineStatus(update, current.id, "rejected")
    } finally {
      update.close()
    }
    return c.json({ ok: true, ...await rebuildCombinedFile(combined) })
  })
})

// Accept every pending redline on a document in one pass: bake them in document
// order, preserve its source/deliverable role, and mark all accepted.
app.post("/matters/:id/redlines/accept-all", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const requested = c.req.query("name") ?? ""
  const name = matterDocumentName(requested) ?? deliverableDocumentName(requested)
  if (!name) return c.json({ error: "The document filename or extension is not supported." }, 400)
  const file = path.join(dir, name)
  return withArtifactMutationLocksAsync(dir, [name, ...runRedlineLock(dir, file), ...runNotesLock(dir), ...runCombinedLock(dir, file), ".dochaus/source-corpus"], async () => {
    if (!existsSync(file) || !existsSync(path.join(dir, ".dochaus", "legal.db")))
      return c.json({ ok: true, accepted: 0 })
    const db = openDb(dir)
    const rows = listPendingRedlines(db, file)
    db.close()
    if (!rows.length) return c.json({ ok: true, accepted: 0 })
    const combined = await prepareCombinedFile(dir, file)
    let redline: PreparedRedline
    try {
      const original = await Bun.file(file).bytes()
      const prior = createHash("sha256").update(original).digest("hex")
      const working = readWorkingDocx(dir,file)
      const state = await redlineState(original,rows)
      const stale = rows.map((row) => redlineStaleReason(row,state,Boolean(working))).find(Boolean)
      if (stale || working && working.source_sha256 !== prior) return c.json({error:stale ?? "The working Word revision is inconsistent. Refresh it through the revision API."},409)
      const baked = documentRevisionBytes(dir,name,Buffer.from(await bake(original,rows,state))).bytes
      redline = await prepareDeliveredRedline(dir,file,baked)
      await writeDocumentRevision(dir,name,baked,rows.map((row) => row.id),prior,rows.some((row) => row.target_json != null))
    } catch (error) {
      console.warn("[redline] batch acceptance refused",diagnosticErrorMetadata(error))
      return c.json({error:"Batch acceptance could not be confirmed. Refresh the document and change list before trying again."},409)
    }
    return c.json({ ok: true, accepted: rows.length, ...await writeDeliveredRedline(redline), ...await recordAcceptedChangesInRunNotes(dir, file, rows), ...await rebuildCombinedFile(combined) })
  })
})

app.post("/matters/:id/redlines/reject-all", async (c) => {
  const dir = matterDir(c.req.param("id"))
  const requested = c.req.query("name") ?? ""
  const name = matterDocumentName(requested) ?? deliverableDocumentName(requested)
  if (!name) return c.json({ error: "The document filename or extension is not supported." }, 400)
  const file = path.join(dir, name)
  return withArtifactMutationLocksAsync(dir, [name, ...rejectCombinedLock(dir, file)], async () => {
    if (!existsSync(path.join(dir, ".dochaus", "legal.db"))) return c.json({ ok: true, rejected: 0 })
    const db = openDb(dir)
    let rows: RedlineRow[]
    try {
      rows = listPendingRedlines(db, file)
    } finally {
      db.close()
    }
    if (!rows.length) return c.json({ ok: true, rejected: 0 })
    // As accept-all does: a rebuild of the saved combined file skipped earlier is made now.
    const combined = await prepareCombinedFile(dir, file, { rejected: true })
    const update = openDb(dir)
    try {
      update.transaction(() => rows.forEach((row) => setRedlineStatus(update, row.id, "rejected")))()
    } finally {
      update.close()
    }
    return c.json({ ok: true, rejected: rows.length, ...await rebuildCombinedFile(combined) })
  })
})

async function writeDocumentRevision(
  directory: string,
  name: string,
  input: Buffer,
  acceptedRedlineIDs: readonly number[] = [],
  priorSha256?: string,
  boundAcceptance = false,
) {
  const { role, bytes } = documentRevisionBytes(directory, name, input)
  const existing = readWorkingDocx(directory,name)
  const working = role === "deliverable" && (boundAcceptance || existing) && name.toLowerCase().endsWith(".docx")
    ? await acceptedWorkingDocxUpdate(directory,name,bytes) : undefined
  const updateRegistry = (registry: ReturnType<typeof openDb>) => {
    acceptedRedlineIDs.forEach((id) => setRedlineStatus(registry,id,"accepted"))
  }
  if (role === "deliverable") return writeDeliverableRevision(directory, name, bytes, { updateRegistry,priorSha256,updateWorkingRevision:working })
  return ingestDocument(directory, name, bytes, { existing: "replace", requireExistingSource: true, updateRegistry })
}

/** The document's role and the bytes writeDocumentRevision stores for `input`:
 * a deliverable docx gains paragraph pointers. Idempotent, so an accept can build
 * what depends on the stored body (its run's redline) before writing it. */
function documentRevisionBytes(directory: string, name: string, input: Buffer) {
  const db = openDb(directory)
  const role = (() => { try { return getDocumentRole(db, path.join(directory, path.basename(name))) } finally { db.close() } })()
  return { role, bytes: role === "deliverable" && name.toLowerCase().endsWith(".docx") ? Buffer.from(withParagraphPointers(input)) : input }
}

// Changes accepted into the body of a completed lean run are appended to that
// run's registered drafting notes (notes-addendum.ts), which otherwise describe
// only the checked draft. The run's delivery summary links its body and notes;
// nothing new is stored. This runs after the body write has committed, inside
// the accept's lock set, which already holds the notes lock (runNotesLock), so
// a rebuild of the saved combined file never pairs the accepted body with the
// notes from before it; the nested lock calls here re-enter the held locks. The
// notes go through the same revision writer as an accepted Word change, so they
// keep their name and a later delivery neither writes over nor forks them.
// Best-effort: the accepted body stands, a skipped addendum is logged and
// returned for the accept response, and the run status still discloses the
// change (changed_after_check). Returns nothing when no addendum applies: the
// accepted file is not a deliverable body of a completed run.
type NotesSkipReason = "notes_not_delivered" | "notes_missing" | "notes_not_deliverable" | "pending_notes_proposals" | "write_failed"
type NotesAddendum = { notes_updated: true } | { notes_updated: false; notes_skip_reason: NotesSkipReason }
async function recordAcceptedChangesInRunNotes(directory: string, body: string, rows: readonly RedlineRow[]): Promise<NotesAddendum | undefined> {
  if (!rows.length) return
  // The server's local calendar date, as a lawyer reads the date of their accept.
  const now = new Date()
  const acceptedAt = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`
  let run: string | undefined
  const skipped = (reason: NotesSkipReason, error?: unknown): NotesAddendum => {
    console.warn("[redline] drafting notes not updated",
      { run, matter_ref: draftingTraceReference(directory), reason, ...(error === undefined ? {} : diagnosticErrorMetadata(error)) })
    return { notes_updated: false, notes_skip_reason: reason }
  }
  try {
    const runBody = deliveredRunBody(directory, body)
    if (!runBody) return
    const delivered = runBody.name
    run = runBody.run
    const notes = runBody.summary.documents.find((row) => row.kind === "drafting_notes")?.deliverable
    if (!notes || notes !== path.basename(notes)) return skipped("notes_not_delivered")
    return await withArtifactMutationLocksAsync(directory, [notes, ".dochaus/source-corpus"], async (): Promise<NotesAddendum> => {
      const file = path.join(directory, notes)
      if (!existsSync(file)) return skipped("notes_missing")
      const db = openDb(directory)
      const [role, pending] = (() => { try { return [getDocumentRole(db, file), listPendingRedlines(db, file)] as const } finally { db.close() } })()
      // A file that has since taken the notes' name, such as a source, is never written.
      if (role !== "deliverable") return skipped("notes_not_deliverable")
      const prior = await Bun.file(file).bytes()
      const next = withParagraphPointers(appendAcceptedChangesToNotes(prior, rows.map((row) => ({
        document: delivered, before: row.old_text, after: row.new_text, reason: row.review_comment || undefined, acceptedAt,
      }))))
      // Never invalidate a proposal on the notes that the lawyer has yet to review.
      if (pending.length) {
        const [before, after] = await Promise.all([redlineState(prior, pending), redlineState(next, pending)])
        const protectedWord = Boolean(readWorkingDocx(directory, file))
        if (pending.some((row) => !redlineStaleReason(row, before, protectedWord) && redlineStaleReason(row, after, true)))
          return skipped("pending_notes_proposals")
      }
      await writeDocumentRevision(directory, notes, Buffer.from(next), [], createHash("sha256").update(prior).digest("hex"), true)
      return { notes_updated: true }
    })
  } catch (error) {
    return skipped("write_failed", error)
  }
}

/** The completed run's delivered notes, as the extra lock an accept takes with
 * its body's so the body write and the notes addendum form one critical section.
 * Read before locking; empty when there are no such notes or they are a link,
 * which the addendum then skips. */
function runNotesLock(directory: string): string[] {
  try {
    const summary = readLeanDocuments(directory)
    if (readLeanState(directory)?.status !== "completed" || summary?.status !== "completed") return []
    const notes = summary.documents.find((row) => row.kind === "drafting_notes")?.deliverable
    return notes && deliverableDocumentName(notes) === notes && readDocumentTarget(directory, notes) ? [notes] : []
  } catch {
    return []
  }
}

/** A completed lean run's delivery of `body`, an accepted file's matter path:
 * the run, its documents summary and that body's summary row (its planned
 * document, title and delivered name), while the file is still a deliverable.
 * Undefined when the accepted file is no delivered body of a completed run. */
function deliveredRunBody(directory: string, body: string) {
  const state = readLeanState(directory)
  const summary = readLeanDocuments(directory)
  if (state?.status !== "completed" || summary?.status !== "completed") return
  const identity = documentRegistryIdentity(body)
  const row = summary.documents.find((row) => row.kind === "body" && row.deliverable === path.basename(row.deliverable ?? "") &&
    documentRegistryIdentity(path.join(directory, row.deliverable!)) === identity)
  const name = row?.deliverable
  if (!row?.document || !name) return
  const registry = openDb(directory)
  try {
    if (getDocumentRole(registry, path.join(directory, name)) !== "deliverable") return
  } finally {
    registry.close()
  }
  return { run: state.owner_session_id, summary, document: row.document, title: row.title, name }
}

/** The delivered redline of a delivered body's document, under its plain registered name. */
function deliveredRedlineName(delivered: NonNullable<ReturnType<typeof deliveredRunBody>>) {
  const name = delivered.summary.documents.find((row) => row.kind === "redline" && row.document === delivered.document && row.title === delivered.title)?.deliverable
  return name && deliverableDocumentName(name) === name ? name : undefined
}

// A change accepted into the body of a completed lean run also rebuilds the
// run's delivered redline of that body (documentRedline: the run's comparison
// baseline against the accepted body, the run's change explanations as Word
// comments), so the viewer and downloads never pair the accepted body with the
// redline from before it. The run's own copy (d<n>/redline.docx) is a stage
// input and stays as it is, so no stage re-runs. The redline is built from the
// body bytes about to be written, before the body commits (prepareDeliveredRedline),
// and written once the body has committed (writeDeliveredRedline): only the
// redline write itself can then fail after the body. Like the notes addendum,
// this runs inside the accept's lock set, which holds the redline's lock
// (runRedlineLock; when that read finds none, each step takes the lock and the
// write checks the redline again), and goes through the same revision writer;
// a revision upload waits on the body's lock. Revisions already
// in the delivered redline keep their dates. Best-effort: the accepted body
// stands, and the accept response says whether the redline was rebuilt or why
// not (logged too). Nothing is returned when the accepted file is no delivered
// body of a completed run, which has no redline to rebuild.
type RedlineSkipReason = "redline_not_delivered" | "redline_missing" | "redline_not_deliverable" | "pending_redline_proposals" | "rebuild_failed" | "write_failed"
type RedlineRebuild = { redline_updated: true } | { redline_updated: false; redline_skip_reason: RedlineSkipReason }
type PreparedRedline = { directory: string; body: string; run: string; name: string; prior: string; bytes: Uint8Array } | RedlineRebuild | undefined
function redlineSkipped(directory: string, run: string | undefined, reason: RedlineSkipReason, error?: unknown): RedlineRebuild {
  console.warn("[redline] delivered redline not rebuilt",
    { run, matter_ref: draftingTraceReference(directory), reason, ...(error === undefined ? {} : diagnosticErrorMetadata(error)) })
  return { redline_updated: false, redline_skip_reason: reason }
}

async function prepareDeliveredRedline(directory: string, body: string, final: Uint8Array): Promise<PreparedRedline> {
  let run: string | undefined
  try {
    const delivered = deliveredRunBody(directory, body)
    if (!delivered) return
    run = delivered.run
    const name = deliveredRedlineName(delivered)
    if (!name) return redlineSkipped(directory, run, "redline_not_delivered")
    return await withArtifactMutationLocksAsync(directory, [name, ".dochaus/source-corpus"], async (): Promise<PreparedRedline> => {
      const blocked = redlineBlocked(directory, name)
      if (blocked) return redlineSkipped(directory, run, blocked)
      const prior = await Bun.file(path.join(directory, name)).bytes()
      const rebuilt = await documentRedline(directory, delivered.document, final)
      const bytes = withStableRevisionDates(rebuilt.bytes, new Date().toISOString(), prior)
      return { directory, body, run: delivered.run, name, prior: createHash("sha256").update(prior).digest("hex"), bytes }
    })
  } catch (error) {
    return redlineSkipped(directory, run, "rebuild_failed", error)
  }
}

/** Writes a prepared redline once the body has committed. The redline's lock may
 * have been released since it was built (runRedlineLock is read before the
 * accept locks), so the name and the redline file are checked again under it:
 * the name must still be the body's delivered redline, the file must still hold
 * the bytes it was built against (the write's prior hash) and take no change
 * proposed in between. */
async function writeDeliveredRedline(prepared: PreparedRedline): Promise<RedlineRebuild | undefined> {
  if (!prepared || "redline_updated" in prepared) return prepared
  const { directory, body, run, name, prior, bytes } = prepared
  try {
    return await withArtifactMutationLocksAsync(directory, [name, ".dochaus/source-corpus"], async (): Promise<RedlineRebuild> => {
      const delivered = deliveredRunBody(directory, body)
      if (!delivered || deliveredRedlineName(delivered) !== name) return redlineSkipped(directory, run, "redline_not_delivered")
      const blocked = redlineBlocked(directory, name)
      if (blocked) return redlineSkipped(directory, run, blocked)
      await writeDocumentRevision(directory, name, Buffer.from(bytes), [], prior)
      return { redline_updated: true }
    })
  } catch (error) {
    return redlineSkipped(directory, run, "write_failed", error)
  }
}

/** Why the delivered redline `name` must be left as it is, read under its lock;
 * undefined when it may be rebuilt. */
function redlineBlocked(directory: string, name: string): RedlineSkipReason | undefined {
  const file = path.join(directory, name)
  if (!existsSync(file)) return "redline_missing"
  const db = openDb(directory)
  const [role, pending] = (() => { try { return [getDocumentRole(db, file), listPendingRedlines(db, file)] as const } finally { db.close() } })()
  // A file that has since taken the redline's name, such as a source, is never written.
  if (role !== "deliverable") return "redline_not_deliverable"
  // The agent's tracked-changes and redline tools can propose changes to the
  // redline file itself; never invalidate one the lawyer has yet to review.
  // Proposals take the file's artifact lock, which this check runs under.
  if (pending.length) return "pending_redline_proposals"
}

/** The delivered redline of a completed run's accepted body, as the extra lock
 * an accept takes with its body's, so the body write and the redline rebuild
 * form one critical section. Read before locking; empty when the accepted file
 * is no such body or its redline is not delivered under a plain name, so the
 * redline write checks its name and file again under its own lock. */
function runRedlineLock(directory: string, body: string): string[] {
  try {
    const delivered = deliveredRunBody(directory, body)
    const name = delivered && deliveredRedlineName(delivered)
    return name && readDocumentTarget(directory, name) ? [name] : []
  } catch {
    return []
  }
}

// A change accepted into a part of a completed lean run's saved combined file
// (registeredCombinedParts: its document's body, or the requested memo's notes,
// which the notes addendum for any delivered body also changes), or a revision
// uploaded in place of a part (/documents/revisions), rebuilds that file from the
// parts as written (buildCombinedFile), so it stays current with them. The
// combined file and both parts are in the accept's or upload's lock set
// (runCombinedLock). Only a file the system wrote and nothing has written since
// is rewritten: the file must hold exactly the bytes recorded by the system's
// last write of it (the delivery's registration or an earlier rebuild), with no
// proposal on it pending: replaceableDeliverable, the rule a run's save over
// the file follows too. A lawyer's accept into the combined file, a revision uploaded in its
// place or a change made outside doc.haus removes or breaks that record, and the
// file is then kept as it is. The rebuild is written with paragraph pointers as
// the run's save gives them (the saved file's where the text is unchanged) and
// recorded, by the bytes written, as the system's own write, so a rebuild
// skipped for any reason (a pending proposal on the combined file, a refused
// build, a failed write) leaves the file the system's and the next accept or
// upload tries again. A reject of a proposal on the combined file or a part
// tries again too (the same prepare and rebuild), so a rebuild held back by the
// proposal is made once it is rejected; as the reject changes no part, it
// writes only when the build differs from the saved file. Under the file's lock the
// rebuild checks again that the run is completed and still registers this
// combined file with these parts, as the redline write does. Best-effort like
// the redline: the accepted body stands, and the response says whether the
// combined file was rebuilt or why not (logged too). Nothing is returned when no
// completed run registered a combined file or the change touched neither part.
type CombinedSkipReason = "combined_missing" | "combined_not_deliverable" | "pending_combined_proposals" | "combined_changed" | "build_refused" | "rebuild_failed" | "write_failed"
type CombinedRebuild = { combined_updated: true } | { combined_updated: false; combined_skip_reason: CombinedSkipReason }
type PreparedCombined = { directory: string; accepted: string; file: string; body: string; notes: string; rejected: boolean; parts?: string; error?: unknown }

function combinedSkipped(directory: string, reason: CombinedSkipReason, detail: Record<string, unknown> = {}): CombinedRebuild {
  console.warn("[redline] combined file not rebuilt", { matter_ref: draftingTraceReference(directory), reason, ...detail })
  return { combined_updated: false, combined_skip_reason: reason }
}

/** The saved combined file the accept of `accepted` (a matter path) may change:
 * a completed run's registered combined row and its parts under plain names,
 * when the accepted file is one of the parts or a delivered body whose notes
 * addendum changes the memo. Undefined otherwise, as for the run's redline and
 * notes (deliveredRunBody, runNotesLock). With `rejected`, for a reject, the
 * combined file itself counts too: rejecting a proposal on it may free a
 * rebuild it held back. */
function registeredCombined(directory: string, accepted: string, rejected = false) {
  const summary = readLeanDocuments(directory)
  if (readLeanState(directory)?.status !== "completed" || summary?.status !== "completed") return
  const parts = registeredCombinedParts(directory, summary.documents)
  if (!parts) return
  const identity = documentRegistryIdentity(accepted)
  const part = [parts.body, parts.notes, ...(rejected ? [parts.file] : [])].some((name) => documentRegistryIdentity(path.join(directory, name)) === identity)
  return part || deliveredRunBody(directory, accepted) ? { file: parts.file, body: parts.body, notes: parts.notes } : undefined
}

/** The saved combined file and its parts, as the extra locks an accept takes so
 * its writes and the combined rebuild form one critical section. Read before
 * locking; empty when the accept changes no saved combined file's part, or a
 * name is a link, so the rebuild checks the file again under its own lock. */
function runCombinedLock(directory: string, accepted: string, rejected = false): string[] {
  try {
    const combined = registeredCombined(directory, accepted, rejected)
    const names = combined ? [combined.file, combined.body, combined.notes] : []
    return names.every((name) => readDocumentTarget(directory, name)) ? names : []
  } catch {
    return []
  }
}

/** A reject's extra locks: the saved combined file and its parts, with the
 * source corpus lock the rebuild takes, when the rejected file is one of them. */
function rejectCombinedLock(directory: string, rejected: string): string[] {
  const names = runCombinedLock(directory, rejected, true)
  return names.length ? [...names, ".dochaus/source-corpus"] : []
}

const sha256Hex = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")

/** Before the accept writes anything: the saved combined file it may rebuild and
 * the sha256 of its parts as they are now, so the rebuild runs only when the
 * accept changed one. With `rejected`, before a reject: the rebuild runs
 * whatever the parts, and writes only a build that differs from the saved file. */
async function prepareCombinedFile(directory: string, accepted: string, options: { rejected?: boolean } = {}): Promise<PreparedCombined | undefined> {
  const rejected = options.rejected ?? false
  let combined: ReturnType<typeof registeredCombined>
  try {
    combined = registeredCombined(directory, accepted, rejected)
  } catch (error) {
    combinedSkipped(directory, "rebuild_failed", diagnosticErrorMetadata(error))
    return
  }
  if (!combined) return
  const { file, body, notes } = combined
  try {
    return await withArtifactMutationLocksAsync(directory, [file, body, notes, ".dochaus/source-corpus"], async (): Promise<PreparedCombined> => {
      // A link is never read or written as the combined file or one of its parts.
      if (![file, body, notes].every((name) => readDocumentTarget(directory, name))) throw new Error("A combined file part is a link, not a matter document.")
      return { directory, accepted, file, body, notes, rejected, parts: (await combinedPartsSha256(directory, body, notes)).key }
    })
  } catch (error) {
    return { directory, accepted, file, body, notes, rejected, error }
  }
}

async function combinedPartsSha256(directory: string, body: string, notes: string) {
  const [memoBytes, bodyBytes] = await Promise.all([Bun.file(path.join(directory, notes)).bytes(), Bun.file(path.join(directory, body)).bytes()])
  return { key: `${sha256Hex(memoBytes)}:${sha256Hex(bodyBytes)}`, memoBytes, bodyBytes }
}

/** After the accept's writes: rebuilds the prepared combined file from its parts
 * as written, when they changed, the run still registers it and the system's
 * last write of it is unchanged. After a reject: the same, when the build
 * differs from the saved file (a rebuild skipped earlier); nothing otherwise,
 * checked first, so a reject that owes no rebuild reports no skip. */
async function rebuildCombinedFile(prepared: PreparedCombined | undefined): Promise<CombinedRebuild | undefined> {
  if (!prepared) return
  const { directory, accepted, file, body, notes } = prepared
  // The parts could not be read before the accept, so whether it changed them is unknown.
  if (!prepared.parts) return combinedSkipped(directory, "rebuild_failed", diagnosticErrorMetadata(prepared.error))
  try {
    return await withArtifactMutationLocksAsync(directory, [file, body, notes, ".dochaus/source-corpus"], async (): Promise<CombinedRebuild | undefined> => {
      if (![file, body, notes].every((name) => readDocumentTarget(directory, name))) throw new Error("A combined file part is a link, not a matter document.")
      const { key, memoBytes, bodyBytes } = await combinedPartsSha256(directory, body, notes)
      if (prepared.parts === key && !prepared.rejected) return
      const target = path.join(directory, file)
      const saved = existsSync(target) ? await Bun.file(target).bytes() : undefined
      const built = buildCombinedFile({ memo: memoBytes, body: bodyBytes })
      // Paragraph pointers as the run's save gives them, the saved file's where the text is unchanged.
      const bytes = saved && !("refused" in built) ? Buffer.from(withParagraphPointers(built.bytes, saved)) : undefined
      // A reject changed no part: a saved file that already holds the build owes
      // no rebuild, so nothing is reported, whatever would hold one back.
      if (prepared.rejected && saved && bytes?.equals(saved)) return
      // The run or its summary may have changed since the accept read them.
      const current = registeredCombined(directory, accepted, prepared.rejected)
      if (current?.file !== file || current.body !== body || current.notes !== notes) return combinedSkipped(directory, "combined_not_deliverable")
      const blocked = combinedBlocked(directory, file)
      if (blocked) return combinedSkipped(directory, blocked)
      // Pending proposals were ruled out above (combinedBlocked), so a refusal here is a change to the file.
      if (!replaceableDeliverable(directory, file)) return combinedSkipped(directory, "combined_changed")
      if ("refused" in built) return combinedSkipped(directory, "build_refused", { refused: built.refused, ...(built.part ? { part: built.part } : {}) })
      // combinedBlocked found the file, read above under this lock.
      if (!saved || !bytes) return combinedSkipped(directory, "combined_missing")
      await writeDeliverableRevision(directory, file, bytes, { priorSha256: sha256Hex(saved), system: true })
      return { combined_updated: true }
    })
  } catch (error) {
    return combinedSkipped(directory, "write_failed", diagnosticErrorMetadata(error))
  }
}

/** Why the saved combined file `name` must be left as it is, read under its
 * lock; undefined when it may be rebuilt. */
function combinedBlocked(directory: string, name: string): CombinedSkipReason | undefined {
  const file = path.join(directory, name)
  if (!existsSync(file)) return "combined_missing"
  const db = openDb(directory)
  const [role, pending] = (() => { try { return [getDocumentRole(db, file), listPendingRedlines(db, file)] as const } finally { db.close() } })()
  if (role !== "deliverable") return "combined_not_deliverable"
  if (pending.length) return "pending_combined_proposals"
  // The lawyer's accepted Word changes into the combined file itself are never written over.
  if (readWorkingDocx(directory, file)) return "combined_changed"
}

function registeredDocumentName(requested: string) {
  return matterDocumentName(requested) ?? deliverableDocumentName(requested)
}

function readDocumentTarget(directory: string, name: string) {
  try {
    return resolveMatterDocumentTarget(directory, name)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) return
    throw error
  }
}

// Every install gets the repo-shipped starter playbooks, so seed them on boot.
// The sample matter (the fictional Aldgate Mills letter) is added only on
// request: POST /sample-matter, called when the lawyer chooses it in setup or
// Settings. Dynamic import keeps seed.ts's heavy docx dependency out of the
// request path's module graph.
const { seedPlaybooks, sampleMatterRoutes } = await import("./seed")
seedPlaybooks()
app.route("/", sampleMatterRoutes)

// Run database compatibility migrations at startup: build/backfill the lexical
// index and refresh vectors, structure and sections when their versions change.
for (const matter of listMatters()) {
  if (!existsSync(path.join(matter.dir, ".dochaus", "legal.db"))) continue
  // One corrupt or externally-locked matter DB must not abort the whole
  // service. A failed step is logged and retried at the next start; the
  // migrations also run again at the matter's next ingest.
  const matterRef = draftingTraceReference(matter.dir)
  const recovery = await recoverDocumentMutations(matter.dir).catch((e) => {
    console.error("matter recovery failed", { matter_ref: matterRef, ...diagnosticErrorMetadata(e) })
  })
  try {
    // The startup open checks the lexical index against the chunks, and only
    // it does, so it runs even when recovery failed: the index derives from
    // the committed chunks alone. The migrations and the artifact heal wait for
    // a recovered matter.
    const db = openDb(matter.dir, { startup: true })
    try {
      if (recovery) {
        const reembedded = await migrateEmbeddings(db)
        if (reembedded) console.log("chunks re-embedded", { count: reembedded, matter_ref: matterRef })
        const restructured = migrateStructure(db)
        if (restructured) console.log("document structure extracted", { count: restructured, matter_ref: matterRef })
        const sectioned = migrateSections(db)
        if (sectioned) console.log("document sections recorded", { count: sectioned, matter_ref: matterRef })
      }
    } finally {
      db.close()
    }
  } catch (e) {
    console.error("matter migration failed", { matter_ref: matterRef, ...diagnosticErrorMetadata(e) })
  }
  if (!recovery) continue
  // The source artifacts derive from the chunks, not from the migrated rows, so
  // a failed migration (rolled back, retried next start) does not stop them
  // being healed.
  try {
    await healSourceArtifacts(matter.dir, Boolean(recovery.sourceRolledBack))
  } catch (e) {
    console.error("source artifact heal failed", { matter_ref: matterRef, ...diagnosticErrorMetadata(e) })
  }
}

// Emails ingested before their attachments were read are re-ingested once,
// after the service is listening, through the usual locked ingest.
setTimeout(async () => {
  for (const matter of listMatters()) {
    if (!existsSync(path.join(matter.dir, ".dochaus", "legal.db"))) continue
    const count = await reingestLegacyEmails(matter.dir).catch((e) => {
      console.error("email re-ingest failed", { matter_ref: draftingTraceReference(matter.dir), ...diagnosticErrorMetadata(e) })
    })
    if (count) console.log("emails re-ingested", { count, matter_ref: draftingTraceReference(matter.dir) })
  }
}, 0)

const preferenceTelemetryRetention = readDraftingTelemetryRetentionPreference()
const envTelemetryRetention = Number(process.env.DOCHAUS_TELEMETRY_RETENTION_DAYS)
const TELEMETRY_RETENTION_DAYS = preferenceTelemetryRetention ?? (Number.isFinite(envTelemetryRetention) && envTelemetryRetention > 0 ? envTelemetryRetention : 90)
if (Number.isFinite(TELEMETRY_RETENTION_DAYS) && TELEMETRY_RETENTION_DAYS > 0) {
  for (const matter of listMatters()) {
    if (!existsSync(path.join(matter.dir, ".dochaus", "telemetry.db"))) continue
    try {
      const db = openTelemetryDb(matter.dir)
      try {
        applyRetentionPolicy(db, TELEMETRY_RETENTION_DAYS)
      } finally {
        db.close()
      }
    } catch (e) {
      console.error("matter telemetry sweep failed", { matter_ref: draftingTraceReference(matter.dir), ...diagnosticErrorMetadata(e) })
    }
  }
}

// Bind loopback by default like the legal engine:
// the service is self-hosted alongside the engine and web app, not exposed
// directly. Front it with a reverse proxy to serve beyond localhost.
const port = Number(process.env.INGEST_PORT ?? 4500)
const hostname = process.env.INGEST_HOST ?? "127.0.0.1"
console.log(`doc.haus ingest service listening on http://${hostname}:${port}`)
export default { port, hostname, fetch: app.fetch }
