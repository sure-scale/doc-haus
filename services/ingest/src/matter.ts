import { createHash, randomUUID } from "node:crypto"
import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync, renameSync, writeFileSync, statSync, rmSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import type { ContentfulStatusCode } from "hono/utils/http-status"
import { withArtifactMutationLocks } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { commitDocumentWrite } from "./atomic-document-write"
import { dataPaths } from "../../engine/src/data-paths"

// A matter is a directory under WORKSPACE_ROOT holding a matter.json plus the
// uploaded documents. This is the same directory the opencode server scopes a
// session to via the x-opencode-directory header.

// Resolved per call so a changed WORKSPACE_ROOT is honoured, not pinned by
// whichever module imported this file first. Every reader of the workspace
// (matters, preferences, playbooks, skills, agents) goes through this.
// WORKSPACE_ROOT is the boot value.
export const workspaceRoot = () => dataPaths(path.resolve(import.meta.dir, "../../.."), process.env, path.join(process.cwd(), "workspace")).workspace
export const WORKSPACE_ROOT = workspaceRoot()

// reference is the firm's own client-matter number (e.g. "2026-0042"), shown in
// the UI. It is display metadata only — the directory id stays the auto slug+uuid.
// jurisdictions are pack codes (e.g. ["EW", "US-NY"]) the engine reads from
// matter.json to steer reasoning and citation style; a matter can span several
// (cross-border deal), so it is a list — see dochaus/jurisdiction/ (issue #18).
// playbook is the single playbook skill name (e.g. "playbook-nda") the matter is
// reviewed against; one matter binds at most one playbook, so it is a string.
export type Matter = {
  id: string
  title: string
  reference?: string
  jurisdictions?: string[]
  playbook?: string
  conflicts?: Conflicts
  dir: string
  created_at: number
}

// The conflicts intake record: proof that a conflicts check happened and how it
// exited — never a judgement about the conflicts themselves. Exactly three
// recordable states: cleared (by whom), pending (owned by whom, optionally due),
// or bypassed (by whom and why — a permanent structured override, never
// auto-removed). Absent means the matter predates the field; downstream behavior
// keys off this record deterministically (dochaus/plugin/legal/policy.ts).
export type Conflicts = {
  status: "cleared" | "pending" | "bypassed"
  by?: string
  owner?: string
  due?: string
  reason?: string
  recorded_at: number
}

function matterFile(dir: string) {
  return path.join(dir, "matter.json")
}

const MATTER_LINEAGE_ARTIFACT = ".dochaus/lineage.jsonl"
const MATTER_ORIGIN_INTENT_ARTIFACT = ".dochaus/matter-origin-mutation.json"
const MATTER_MUTATION_ARTIFACTS = ["matter.json", MATTER_LINEAGE_ARTIFACT, MATTER_ORIGIN_INTENT_ARTIFACT]

type MatterMutationKind = "create" | "rename" | "conflicts"

type MatterOriginIntent = {
  version: 1
  artifact: "matter.json"
  mutation_id: string
  correlation_id: string
  mutation_kind: MatterMutationKind
  entry_id: string
  ts: string
  content_sha256: string
  bytes: number
  supersedes: string | null
}

// A matter id is the auto-generated slug+uuid from createMatter, so it is always
// [a-z0-9-]. Reject anything else before it reaches a filesystem path: the id
// arrives straight from a URL param, and matterDir is the single chokepoint every
// route resolves through, so guarding here blocks path traversal (../, absolute
// paths) out of WORKSPACE_ROOT across the whole service.
export function matterDir(id: string) {
  if (!/^[a-z0-9][a-z0-9-]*$/.test(id)) throw new Error(`invalid matter id: ${id}`)
  return path.join(workspaceRoot(), id)
}

export function listMatters(): Matter[] {
  const root = workspaceRoot()
  mkdirSync(root, { recursive: true })
  return readdirSync(root)
    .map((name) => path.join(root, name))
    // One hand-edited or partially written metadata file must not take down the
    // matter list, startup maintenance, or the drafting recovery sweep for
    // every other matter. Invalid records stay isolated to their directory.
    .flatMap((dir) => {
      const current = readListedMatter(dir)
      return current ? [current] : []
    })
    .sort((a, b) => a.created_at - b.created_at)
}

function readListedMatter(dir: string): Matter | undefined {
  try {
    recoverMatterOriginMutation(dir)
    if (!statSync(dir).isDirectory() || !existsSync(matterFile(dir))) return
    const parsed = JSON.parse(readFileSync(matterFile(dir), "utf8")) as unknown
    if (!isRecord(parsed)) return
    if (
      typeof parsed.id !== "string" ||
      parsed.id !== path.basename(dir) ||
      typeof parsed.title !== "string" ||
      typeof parsed.created_at !== "number" ||
      !Number.isFinite(parsed.created_at)
    ) return
    return { ...parsed, dir } as Matter
  } catch {
    return
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

export function createMatter(title: string, reference?: string, jurisdictions?: string[], playbook?: string): Matter {
  const slug = title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")
  const id = `${slug}-${crypto.randomUUID().slice(0, 6)}`
  const dir = matterDir(id)
  mkdirSync(dir, { recursive: true })
  const matter: Matter = { id, title, reference, jurisdictions, playbook, dir, created_at: Date.now() }
  return withMatterMutationLocks(dir, () => persistMatterLocked(matter, "create"))
}

// Rename keeps the directory id stable (it backs every session and document
// path); only the display title/reference/jurisdictions/playbook in matter.json change.
export function renameMatter(id: string, title: string, reference?: string, jurisdictions?: string[], playbook?: string): Matter {
  return mutateMatter(id, "rename", (matter) => ({ ...matter, title, reference, jurisdictions, playbook }))
}

// Records a conflicts-check outcome on the matter. Each exit carries the fields
// that make the record auditable: cleared names who ran the check, pending names
// who owns finishing it, bypassed names who decided and why. Anything else is
// rejected so matter.json never holds a half-recorded check.
export function recordConflicts(
  id: string,
  input: { status: Conflicts["status"]; by?: string; owner?: string; due?: string; reason?: string },
): Matter {
  if (!["cleared", "pending", "bypassed"].includes(input.status))
    throw new Error(`conflicts status must be "cleared", "pending", or "bypassed", got "${input.status}"`)
  if (input.status === "cleared" && !input.by)
    throw new Error('conflicts "cleared" requires "by" — who performed the check')
  if (input.status === "pending" && !input.owner)
    throw new Error('conflicts "pending" requires "owner" — who is responsible for completing the check')
  if (input.status === "pending" && input.due !== undefined && Number.isNaN(Date.parse(input.due)))
    throw new Error(`conflicts "due" must be an ISO date, got "${input.due}"`)
  if (input.status === "bypassed" && (!input.by || !input.reason))
    throw new Error('conflicts "bypassed" requires "by" and "reason" — a documented bypass names who decided and why')
  return mutateMatter(id, "conflicts", (matter) => ({
    ...matter,
    conflicts: {
      status: input.status,
      by: input.by,
      owner: input.owner,
      due: input.due,
      reason: input.reason,
      recorded_at: Date.now(),
    },
  }))
}

export function deleteMatter(id: string) {
  rmSync(matterDir(id), { recursive: true, force: true })
}

export function getMatter(id: string): Matter {
  const dir = matterDir(id)
  recoverMatterOriginMutation(dir)
  return JSON.parse(readFileSync(matterFile(dir), "utf8")) as Matter
}

function mutateMatter(id: string, kind: Exclude<MatterMutationKind, "create">, update: (matter: Matter) => Matter) {
  const dir = matterDir(id)
  return withMatterMutationLocks(dir, () => {
    recoverMatterOriginMutationLocked(dir)
    return persistMatterLocked(update(JSON.parse(readFileSync(matterFile(dir), "utf8")) as Matter), kind)
  })
}

function persistMatterLocked(matter: Matter, kind: MatterMutationKind) {
  recoverMatterOriginMutationLocked(matter.dir)
  const content = Buffer.from(JSON.stringify(matter, null, 2))
  const state = contentState(content)
  const intent: MatterOriginIntent = {
    version: 1,
    artifact: "matter.json",
    mutation_id: randomUUID(),
    correlation_id: `matter-${kind}-${randomUUID()}`,
    mutation_kind: kind,
    entry_id: randomUUID(),
    ts: new Date().toISOString(),
    content_sha256: state.sha256,
    bytes: state.bytes,
    supersedes: lastMatterLineageEntryID(matter.dir),
  }
  writeMatterOriginIntent(matter.dir, intent)
  commitDocumentWrite(matterFile(matter.dir), content, () => appendMatterOriginReceiptLocked(matter.dir, intent))
  rmSync(matterOriginIntentFile(matter.dir), { force: true })
  return matter
}

// A fixed, hash-only intent closes the process-crash gap between replacing
// matter.json and appending its receipt. Existing matters without this intent
// are deliberately left alone: recovery is not a metadata backfill.
function recoverMatterOriginMutation(dir: string) {
  if (!existsSync(matterOriginIntentFile(dir))) return false
  return withMatterMutationLocks(dir, () => recoverMatterOriginMutationLocked(dir))
}

function recoverMatterOriginMutationLocked(dir: string) {
  const file = matterOriginIntentFile(dir)
  if (!existsSync(file)) return false
  const intent = readMatterOriginIntent(file)
  const current = fileState(matterFile(dir))
  if (!current || current.sha256 !== intent.content_sha256 || current.bytes !== intent.bytes) {
    rmSync(file, { force: true })
    return false
  }
  appendMatterOriginReceiptLocked(dir, intent)
  rmSync(file, { force: true })
  return true
}

function appendMatterOriginReceiptLocked(dir: string, intent: MatterOriginIntent) {
  const current = fileState(matterFile(dir))
  if (!current || current.sha256 !== intent.content_sha256 || current.bytes !== intent.bytes)
    throw new Error(`Matter mutation ${intent.mutation_id} does not match the current matter.json bytes.`)
  const expected = matterOriginReceipt(intent)
  const existing = lineageRows(dir).find(
    (row) => isRecord(row.provenance) && row.provenance.matter_mutation_id === intent.mutation_id,
  )
  if (existing) {
    if (!sameMatterOriginReceipt(existing, expected))
      throw new Error(`Matter mutation ${intent.mutation_id} has conflicting lineage evidence.`)
    return
  }
  const file = path.join(dir, MATTER_LINEAGE_ARTIFACT)
  mkdirSync(path.dirname(file), { recursive: true })
  const prior = existsSync(file) ? readFileSync(file, "utf8") : ""
  appendFileSync(file, `${prior.length && !prior.endsWith("\n") ? "\n" : ""}${JSON.stringify(expected)}\n`)
}

function matterOriginReceipt(intent: MatterOriginIntent) {
  return {
    entry_id: intent.entry_id,
    ts: intent.ts,
    tier: "PERMANENT",
    artifact: "matter.json",
    op: intent.mutation_kind === "create" ? "write" : "edit",
    run: `ingest:${intent.correlation_id}`,
    stage: null,
    subagent: "ingest",
    content_sha256: intent.content_sha256,
    bytes: intent.bytes,
    supersedes: intent.supersedes,
    provenance: {
      origin: "matter-metadata",
      matter_mutation_id: intent.mutation_id,
      correlation_id: intent.correlation_id,
      mutation_kind: intent.mutation_kind,
    },
  }
}

function sameMatterOriginReceipt(current: Record<string, unknown>, expected: ReturnType<typeof matterOriginReceipt>) {
  return (
    current.entry_id === expected.entry_id &&
    current.ts === expected.ts &&
    current.tier === expected.tier &&
    current.artifact === expected.artifact &&
    current.op === expected.op &&
    current.run === expected.run &&
    current.stage === expected.stage &&
    current.subagent === expected.subagent &&
    current.content_sha256 === expected.content_sha256 &&
    current.bytes === expected.bytes &&
    current.supersedes === expected.supersedes &&
    isRecord(current.provenance) &&
    current.provenance.origin === expected.provenance.origin &&
    current.provenance.matter_mutation_id === expected.provenance.matter_mutation_id &&
    current.provenance.correlation_id === expected.provenance.correlation_id &&
    current.provenance.mutation_kind === expected.provenance.mutation_kind
  )
}

function lastMatterLineageEntryID(dir: string) {
  return lineageRows(dir).reduce<string | null>(
    (entry, row) => row.artifact === "matter.json" && typeof row.entry_id === "string" ? row.entry_id : entry,
    null,
  )
}

function lineageRows(dir: string) {
  const file = path.join(dir, MATTER_LINEAGE_ARTIFACT)
  if (!existsSync(file)) return []
  return readFileSync(file, "utf8")
    .split(/\r?\n/)
    .flatMap((line) => {
      if (!line.trim()) return []
      try {
        const parsed = JSON.parse(line) as unknown
        return isRecord(parsed) ? [parsed] : []
      } catch {
        return []
      }
    })
}

function writeMatterOriginIntent(dir: string, intent: MatterOriginIntent) {
  const file = matterOriginIntentFile(dir)
  const temporary = `${file}.${process.pid}.${randomUUID()}.tmp`
  mkdirSync(path.dirname(file), { recursive: true })
  writeFileSync(temporary, `${JSON.stringify(intent, null, 2)}\n`, { flag: "wx" })
  try {
    renameSync(temporary, file)
  } finally {
    rmSync(temporary, { force: true })
  }
}

function readMatterOriginIntent(file: string): MatterOriginIntent {
  const parsed = JSON.parse(readFileSync(file, "utf8")) as unknown
  if (
    !isRecord(parsed) ||
    parsed.version !== 1 ||
    parsed.artifact !== "matter.json" ||
    typeof parsed.mutation_id !== "string" ||
    !parsed.mutation_id ||
    typeof parsed.correlation_id !== "string" ||
    !parsed.correlation_id ||
    !["create", "rename", "conflicts"].includes(String(parsed.mutation_kind)) ||
    typeof parsed.entry_id !== "string" ||
    !parsed.entry_id ||
    typeof parsed.ts !== "string" ||
    !/^[a-f0-9]{64}$/.test(String(parsed.content_sha256)) ||
    typeof parsed.bytes !== "number" ||
    !Number.isSafeInteger(parsed.bytes) ||
    parsed.bytes < 0 ||
    !(parsed.supersedes === null || typeof parsed.supersedes === "string")
  ) throw new Error(`Invalid matter origin mutation intent at ${file}.`)
  return parsed as MatterOriginIntent
}

function matterOriginIntentFile(dir: string) {
  return path.join(dir, MATTER_ORIGIN_INTENT_ARTIFACT)
}

function withMatterMutationLocks<T>(dir: string, fn: () => T) {
  return withArtifactMutationLocks(dir, MATTER_MUTATION_ARTIFACTS, fn)
}

function fileState(file: string) {
  if (!existsSync(file)) return null
  return contentState(readFileSync(file))
}

function contentState(content: Buffer) {
  return { sha256: createHash("sha256").update(content).digest("hex"), bytes: content.byteLength }
}

// The jurisdiction packs available to assign to a matter, read from the dochaus
// config layer (dochaus/jurisdiction/<code>/profile.json). The web app fetches
// this to offer the choices; the engine consumes the full pack (issue #18).
// Resolved relative to this module like the template seed, so it does not depend
// on the ingest process cwd.
const JURISDICTION_DIR = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "..",
  "..",
  "dochaus",
  "jurisdiction",
)

export function listJurisdictions() {
  if (!existsSync(JURISDICTION_DIR)) return []
  return readdirSync(JURISDICTION_DIR)
    .map((code) => path.join(JURISDICTION_DIR, code, "profile.json"))
    .filter((file) => existsSync(file))
    .map((file) => JSON.parse(readFileSync(file, "utf8")) as { code: string; name: string; citationStyle: string })
    .sort((a, b) => a.name.localeCompare(b.name))
}

// Playbooks are opencode skills named playbook-*. Two sources feed the list: the
// repo-shipped skills under dochaus/skill/ (resolved relative to this module like
// the jurisdiction packs, so it is cwd-independent; read-only), and the imported
// ones the ingest service writes under WORKSPACE_ROOT/.playbooks. Each SKILL.md
// carries the name and description (plus optional version/last-reviewed) in its
// leading frontmatter block. The web app reads this to offer the playbook choices
// for a matter; the engine binds one via matter.json.
const SKILL_DIR = path.join(path.dirname(JURISDICTION_DIR), "skill")
export const PLAYBOOK_DIR = () => path.join(workspaceRoot(), ".playbooks")

export type Playbook = {
  name: string
  description: string
  content: string
  version: string
  last_reviewed: string
  builtin: boolean
  updated_at: number
  matters: string[]
}

export class PlaybookError extends Error {
  status: ContentfulStatusCode
  constructor(message: string, status: ContentfulStatusCode) {
    super(message)
    this.status = status
  }
}

export function listPlaybooks(): Playbook[] {
  // Bound matters are surfaced per playbook so the library shows where each one
  // is in use and deletePlaybook can refuse to orphan a binding.
  const matters = listMatters()
  return [
    { dir: SKILL_DIR, builtin: true },
    { dir: PLAYBOOK_DIR(), builtin: false },
  ]
    .filter((source) => existsSync(source.dir))
    .flatMap((source) =>
      readdirSync(source.dir)
        .map((entry) => path.join(source.dir, entry, "SKILL.md"))
        .filter((file) => existsSync(file))
        .map((file) => ({
          ...parsePlaybookMarkdown(readFileSync(file, "utf8")),
          builtin: source.builtin,
          updated_at: statSync(file).mtimeMs,
        })),
    )
    .filter((p) => p.name.startsWith("playbook-"))
    .map((p) => ({ ...p, matters: matters.filter((m) => m.playbook === p.name).map((m) => m.title) }))
    .sort((a, b) => a.name.localeCompare(b.name))
}

// SKILL.md = frontmatter + markdown body, same shape skill.ts owns for skills.
// Descriptions are JSON-quoted on write, so a quoted value is JSON-parsed back.
function parsePlaybookMarkdown(source: string) {
  const lines = source.split("\n")
  const end = lines[0]?.trim() === "---" ? lines.findIndex((l, i) => i > 0 && l.trim() === "---") : -1
  const read = (key: string) => {
    const line = lines.slice(1, end === -1 ? 0 : end).find((l) => l.startsWith(`${key}:`))
    if (!line) return ""
    const raw = line.slice(key.length + 1).trim()
    // Quoted values are JSON-parsed back; a legacy or hand-edited file may carry
    // a quote-prefixed value that is not valid JSON, so fall back to the raw
    // value (outer quotes stripped) rather than failing the whole playbook list.
    if (raw.startsWith('"')) {
      try {
        return JSON.parse(raw) as string
      } catch {
        return raw.replace(/^"|"$/g, "")
      }
    }
    return raw
  }
  return {
    name: read("name"),
    description: read("description"),
    version: read("version"),
    last_reviewed: read("last-reviewed"),
    content: end === -1 ? source.trim() : lines.slice(end + 1).join("\n").trim(),
  }
}

function renderPlaybookMarkdown(playbook: { name: string; description: string; version: string; last_reviewed: string; content: string }) {
  // Description is JSON-quoted so a colon or newline cannot inject YAML keys
  // (same posture as skill.ts renderSkillMarkdown). version/last-reviewed are
  // carried through when present so an update never strips them.
  const frontmatter = [
    `name: ${playbook.name}`,
    `description: ${JSON.stringify(playbook.description)}`,
    playbook.version && `version: ${JSON.stringify(playbook.version)}`,
    playbook.last_reviewed && `last-reviewed: ${JSON.stringify(playbook.last_reviewed)}`,
  ].filter((line): line is string => Boolean(line))
  return `---\n${frontmatter.join("\n")}\n---\n\n${playbook.content.trim()}\n`
}

// True keep-semantics: omitted fields retain their current values — the server
// owns the file, so it backfills from the existing SKILL.md rather than trusting
// callers to re-send everything. Only imported playbooks in WORKSPACE_ROOT/
// .playbooks are writable; repo-shipped ones are read-only.
export function updatePlaybook(name: string, input: { description?: string; content?: string }): Playbook {
  // basename() once, used for both the lookup and the filesystem path, so a
  // URL-decoded ../ in the name can never write outside PLAYBOOK_DIR (same
  // posture as the POST /playbooks handler).
  const safe = path.basename(name)
  const current = requireImportedPlaybook(safe)
  if (input.description === undefined && input.content === undefined)
    throw new PlaybookError("Nothing to update — supply description, content, or both", 400)
  if (input.content !== undefined && !input.content.trim()) throw new PlaybookError("content must be non-empty", 400)
  const updated = {
    ...current,
    description: input.description ?? current.description,
    content: input.content?.trim() ?? current.content,
  }
  const file = path.join(PLAYBOOK_DIR(), safe, "SKILL.md")
  writeFileSync(file, renderPlaybookMarkdown(updated))
  return { ...updated, updated_at: statSync(file).mtimeMs }
}

// Deleting a playbook a matter is still bound to would leave that matter's
// playbook setting dangling, so it is refused until the bindings are cleared.
export function deletePlaybook(name: string) {
  // basename() once for both the guard and the rmSync path — see updatePlaybook.
  const safe = path.basename(name)
  const current = requireImportedPlaybook(safe)
  if (current.matters.length)
    throw new PlaybookError(
      `playbook "${safe}" is bound to ${current.matters.length} matter(s): ${current.matters.join(", ")}. Unbind it from each matter before deleting.`,
      409,
    )
  rmSync(path.join(PLAYBOOK_DIR(), safe), { recursive: true, force: true })
}

function requireImportedPlaybook(name: string) {
  const playbook = listPlaybooks().find((p) => p.name === name)
  if (!playbook) throw new PlaybookError(`playbook "${name}" not found`, 404)
  if (playbook.builtin) throw new PlaybookError(`playbook "${name}" is repo-shipped and read-only`, 403)
  return playbook
}
