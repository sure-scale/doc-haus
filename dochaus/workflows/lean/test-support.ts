import { Database } from "bun:sqlite"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { withParagraphPointers } from "../../lib/drafting-amendment"
import { recordSystemDocx, replaceableDeliverable } from "../../lib/system-docx"
import { readWorkingDocx } from "../../lib/working-docx"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { LEAN_STAGE_AGENTS, type LeanDeps, type StageSessionPort } from "./stages"
import { initLeanState } from "./state"

// Shared fixtures for the lean drafting tests: a matter with one indexed source
// and a scripted model. Only the model is scripted; replies are keyed by stage
// so they follow the stage table's agents, and every caller drives the real
// coordinator.

export const SOURCE = "supply-terms.docx"
export const SOURCE_TEXT = "The supplier shall deliver the goods within thirty days of the order date. Payment is due on delivery."
export const DRAFT = "# Supply note\n\n## Delivery\n\nThe supplier shall deliver the goods.\n\n## Payment\n\nPayment is due on delivery.\n"
export const SUPPORT = [{ source_id: SOURCE, quote: "within thirty days of the order date" }]
export const AGENT = LEAN_STAGE_AGENTS

export type Reply = (prompt: string, index: number, sessionID?: string) => string

const directories: string[] = []

export function cleanupMatters() {
  directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true }))
}

export function matter(sourceText = SOURCE_TEXT) {
  const directory = createMatter("lean-runner-", sourceText)
  directories.push(directory)
  initLeanState(directory, { request: "Draft a short supply note from the supply terms.", ownerSessionID: "ses_owner", model: { providerID: "test", modelID: "scripted" } })
  return directory
}

/** The current plan, now also naming a requested memo, as an outline does. */
export function requestMemo(directory: string) {
  const plan = readLeanPlan(directory)!
  writeLeanPlan(directory, plan.documents.map(({ id, ...document }) => document), plan.basis, plan.acting_for, { name: "cover memo", file: "cover memo.docx" }, plan.posture)
}

export function scriptedModel(replies: Record<string, Reply>) {
  const prompts: { agent: string; sessionID: string; text: string }[] = []
  const sessions: Parameters<StageSessionPort["createSession"]>[0][] = []
  // Kept across prompts, so a retry answers as the job its session asked.
  const actRecord = actRecordJobs({})
  const port: StageSessionPort = {
    async createSession(input) {
      sessions.push(input)
      return `ses_child${sessions.length}`
    },
    async prompt(input) {
      prompts.push(input)
      const preparation: Record<string, Reply> = {
        // Ordinary drafting regression fixtures assert no source-derived items.
        // Dedicated checklist lifecycle fixtures provide positive items/checks.
        "lean-checklist-extract": () => JSON.stringify({ items: [] }),
        // The act record lists no act and no deliverable; its WHO and HOW jobs,
        // asked only of listed acts, leave each act's parties and instrument
        // unsettled with a reason.
        [AGENT.acts]: actRecord,
        // The outline's files job finds no lawyer instruction on the files to deliver.
        [AGENT.files]: () => JSON.stringify({ stated: false, files: [] }),
        // The plan rejects every listed review finding: the scripted changes
        // stage amends from the findings it is shown, whatever the plan says.
        [AGENT["change-plan"]]: prompt => {
          const resolves = listedFindingIDs(prompt)
          return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [], findings: resolves.length ? [{
            disposition: "rejected", kind: "suggestion", detail: "The scripted plan adjudicates the review findings.", reason: "The scripted changes stage decides the amendments.",
            documents: [/^- (d\d+) "[^"]*" \(this document\)$/m.exec(prompt)?.[1] ?? "d1"], sources: [], limitations: ["Scripted fixture."], resolves,
          }] : [] })
        },
        [AGENT["change-prepare"]]: () => JSON.stringify({ amendments: [], questions: [] }),
        [AGENT["change-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [] }),
        [AGENT.discrepancy]: () => JSON.stringify({ entries: [], limitations: [] }),
        [AGENT.spine]: () => JSON.stringify({ entries: [], limitations: [] }),
        [AGENT["review-independent"]]: () => NO_FINDINGS,
        [AGENT["review-wording"]]: () => NO_FINDINGS,
        [AGENT["review-client-disadvantage"]]: () => NO_FINDINGS,
        [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [] }),
        [AGENT.intake]: () => JSON.stringify({ ready: true, jurisdiction: "No external legal determination needed for this supplied-source fixture", scope: "Reproduce supplied terms only" }),
        [AGENT.research]: () => JSON.stringify({ status: "not-needed", findings: [], limitations: ["This fixture requests a note from supplied terms only."] }),
        [AGENT["research-review"]]: () => JSON.stringify({ approved: [], rejected: [], limitations: [], gaps: [] }),
        // Every asked question is advised; the fixture's advisor recommends nothing.
        [AGENT.advise]: () => JSON.stringify({ unavailable: { reason: "The scripted advisor recommends nothing." } }),
        // Every answer in the lawyer's words is read as their own answer.
        [AGENT.answer]: () => JSON.stringify({ class: "answered" }),
        // Every set of answers can be followed together.
        [AGENT.reconcile]: () => JSON.stringify({ conflicts: [] }),
      }
      const reply = replies[input.agent] ?? preparation[input.agent]
      if (!reply) throw new Error(`No scripted reply for ${input.agent}.`)
      return reply(input.text, prompts.filter(prompt => prompt.agent === input.agent).length - 1, input.sessionID)
    },
  }
  const count = (agent: string) => prompts.filter(prompt => prompt.agent === agent).length
  return { port, prompts, sessions, count, counts: () => [AGENT.draft, AGENT.review, AGENT.changes, AGENT.finalize].map(count) }
}

export type ActRecordJob = "what" | "deliverables" | "who" | "how"

/** Which act-record job a first prompt asks, by the reply shape its contract gives. */
export function actRecordJob(prompt: string): ActRecordJob {
  return prompt.includes('{"index":0,"how":{') ? "how" : prompt.includes('{"index":0,"who":{') ? "who" : prompt.includes('{"deliverables":[') ? "deliverables" : "what"
}

/** The act-record agent scripted per job: each job's session is told apart by
 * its first prompt, and a retry in the same session answers as that job. A
 * job without its own reply gets the default: WHAT lists no act, DELIVERABLES
 * none, and WHO and HOW leave every act they are shown unsettled. */
export function actRecordJobs(replies: Partial<Record<ActRecordJob, Reply>>): Reply {
  const jobs = new Map<string | undefined, ActRecordJob>()
  const unsettled = (prompt: string, field: "who" | "how") => JSON.stringify({ acts: listedActs(prompt).map(({ index }) => ({ index, [field]: field === "who"
    ? { authorised_by: [], performed_by: [], basis: [], reason: "The scripted record does not settle the parties." }
    : { instrument: null, executed_by: [], requires: [], basis: [], reason: "The scripted record does not settle the instrument." } })) })
  const defaults: Record<ActRecordJob, Reply> = {
    what: () => JSON.stringify({ acts: [] }),
    deliverables: () => JSON.stringify({ deliverables: [] }),
    who: prompt => unsettled(prompt, "who"),
    how: prompt => unsettled(prompt, "how"),
  }
  return (prompt, index, sessionID) => {
    if (!jobs.has(sessionID)) jobs.set(sessionID, actRecordJob(prompt))
    const job = jobs.get(sessionID)!
    return (replies[job] ?? defaults[job])(prompt, index, sessionID)
  }
}

/** The acts a coverage placement or act-record WHO or HOW prompt lists, by index. */
export function listedActs(prompt: string) {
  const json = /<acts>\n([\s\S]*?)\n<\/acts>/.exec(prompt)?.[1]
  if (!json) throw new Error("The prompt lists no acts.")
  return JSON.parse(json) as { index: number; act: string; source: string; quote: string; parties: string[] }[]
}

/** One act the scripted act record lists: what is to be done, the words that
 * direct it, their source (the request when not given) and the parties whose
 * approval authorises it ([] when the record does not settle them). */
export type ScriptedAct = { act: string; quote: string; source?: string; parties: string[] }

/** The act-record agent listing `acts` in every WHAT sample, with DELIVERABLES
 * listing none. WHO names each act's parties as those whose approval authorises
 * it, on the act's own words, and HOW has them execute the instrument it leaves
 * unsettled, so no signer is flagged; an act with no parties is left unsettled
 * by both. Given a function, each job reads the acts from its own prompt, as an
 * answer the instructions now hold can change them. Its jobs are told apart by
 * session, so each scripted model needs its own. */
export function actRecordOf(acts: ScriptedAct[] | ((prompt: string) => ScriptedAct[])): Reply {
  const listing = (prompt: string) => (typeof acts === "function" ? acts(prompt) : acts).map(({ act, quote, source = "request", parties }) => ({ act, source, quote, parties }))
  const described = (prompt: string, field: "who" | "how") => JSON.stringify({ acts: listedActs(prompt).map(listed => {
    const parties = listing(prompt).find(act => act.act === listed.act && act.source === listed.source && act.quote === listed.quote)?.parties ?? []
    const basis = [{ source: listed.source, quote: listed.quote }]
    return { index: listed.index, [field]: field === "who"
      ? parties.length ? { authorised_by: parties, performed_by: [], basis, reason: "The scripted record does not settle who carries the act out." }
        : { authorised_by: [], performed_by: [], basis: [], reason: "The scripted record does not settle the parties." }
      : { instrument: null, executed_by: parties, requires: [], basis: parties.length ? basis : [], reason: "The scripted record does not settle the instrument." } }
  }) })
  return actRecordJobs({
    what: prompt => JSON.stringify({ acts: listing(prompt).map(({ parties, ...act }) => act) }),
    who: prompt => described(prompt, "who"),
    how: prompt => described(prompt, "how"),
  })
}

/** The coverage placement job scripted from one reply that gives each act with
 * its document: it places each act it is shown where that reply put it. A reply
 * that is not such a list is returned unchanged. The reply reads the job's own
 * prompt and count. A retry arrives in the same session without the acts, which
 * the session already holds. */
export function placementJob(reply: Reply): Record<string, Reply> {
  const shown = new Map<string | undefined, ReturnType<typeof listedActs>>()
  const placed = (text: string) => {
    try {
      const value = JSON.parse(text)
      if (value && Array.isArray(value.acts)) return value.acts as Record<string, unknown>[]
    } catch {}
  }
  return {
    [AGENT.coverage]: (prompt, index, sessionID) => {
      const text = reply(prompt, index, sessionID)
      const acts = placed(text)
      if (!acts) return text
      const listing = prompt.includes("<acts>\n") ? listedActs(prompt) : shown.get(sessionID)
      if (!listing) throw new Error("The placement session was shown no acts.")
      shown.set(sessionID, listing)
      return JSON.stringify({ placements: listing.flatMap(listed => acts
        .filter(act => act.act === listed.act && act.source === listed.source && act.quote === listed.quote)
        .map(act => ({ act: listed.index, document: act.document ?? null, ...(act.excluded_by !== undefined ? { excluded_by: act.excluded_by } : {}) }))) })
    },
  }
}

/** The draft agent for whole Markdown documents, in drafting order: their "# "
 * titles and "## " parts answer the outline prompt (and any retry of it), after
 * an "Acting for: not stated" line, then each part prompt gets the text of the
 * part it names in the document it names. A "Form:" or "Headings:" line between
 * a title and its first part is an outline line of that document, which is
 * otherwise planned as "Form: other" with "Headings: shown", its "## " parts
 * printed as its headings. Each assembled draft is then the document itself,
 * without those lines. */
export function drafter(...documents: string[]): Reply {
  const planned = documents.map(document => {
    const [head = "", ...parts] = document.split(/^(?=## )/m)
    const lines = head.split("\n").filter(line => /^(?:form|headings)\s*:/i.test(line))
    return {
      title: /^# (.+)$/m.exec(document)?.[1],
      outline: lines.some(line => /^form\s*:/i.test(line)) ? lines : ["Form: other", ...(lines.length ? [] : ["Headings: shown"]), ...lines],
      parts: parts.map(part => part.trim()),
    }
  })
  if (!planned.length || planned.some(row => !row.title || !row.parts.length))
    throw new Error('A scripted draft needs a "# " title and at least one "## " part.')
  return prompt => {
    const named = /^Draft part \d+ of \d+ of "([^"]*)": (.+)$/m.exec(prompt)
    if (!named) return ["Acting for: not stated", ...planned.flatMap(row => [`# ${row.title}`, ...row.outline, ...row.parts.map(part => part.split("\n")[0]!)])].join("\n")
    const part = planned.find(row => row.title === named[1])?.parts.find(row => row.split("\n")[0] === `## ${named[2]}`)
    if (!part) throw new Error(`The scripted draft has no part "${named[2]}" in "${named[1]}".`)
    return part
  }
}

export function pidOf(prompt: string, text: string) {
  const pid = new RegExp(`\\[(p[0-9a-f]{12})\\] ${text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`).exec(prompt)?.[1]
  if (!pid) throw new Error(`The prompt does not list a paragraph starting "${text}".`)
  return pid
}

/** A retry arrives in the same session and does not repeat the draft's
 * paragraph list, which the session already holds. The wrapped reply keeps
 * reading the last prompt that listed the paragraphs. */
export function sameSession(reply: Reply): Reply {
  const seen = { prompt: "" }
  return (prompt, index) => {
    seen.prompt = /^\[p[0-9a-f]{12}\] /m.test(prompt) ? prompt : seen.prompt
    return reply(seen.prompt, index)
  }
}

/** A review reply with no findings. */
export const NO_FINDINGS = JSON.stringify({ findings: [] })

export const review: Reply = prompt => JSON.stringify({ findings: [{
  pid: pidOf(prompt, "The supplier shall deliver the goods."), origin: "source_conflict", balance: "neutral", finding: "The delivery period is missing.", sources: SUPPORT,
}] })

/** The review finding ids a change plan prompt lists for adjudication. */
export function listedFindingIDs(prompt: string) {
  const block = /<review_findings>\n([\s\S]*?)\n<\/review_findings>/.exec(prompt)?.[1] ?? "[]"
  return (JSON.parse(block) as { id: string }[]).map(row => row.id)
}

export const changes: Reply = prompt => JSON.stringify({
  amendments: [{
    pid: pidOf(prompt, "The supplier shall deliver the goods."),
    updated: "The supplier shall deliver the goods within thirty days of the order date.",
    reason: "States the delivery period from the supply terms.",
    sources: SUPPORT,
  }],
  questions: [],
})

export const noChanges: Reply = () => JSON.stringify({ amendments: [], questions: [] })

// Discloses every open item the coordinator lists, in every document's block,
// and every delegated question, as the drafting-notes contract asks: in words,
// each ending with its id in a trailing "[ref: <id>]".
export const draftingNotes: Reply = prompt => {
  const blocks = [...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].map(match => match[1]!)
  const delegated = /<delegated_decisions>\n([\s\S]*?)\n<\/delegated_decisions>/.exec(prompt)?.[1] ?? ""
  const ids = [...blocks.flatMap(block => [...block.matchAll(/^- (\S+):/gm)].map(match => match[1])), ...[...delegated.matchAll(/^- \[(\S+)\] Q:/gm)].map(match => match[1])]
  return ["# Drafting notes", "", "The draft states the delivery period from the supply terms.", "",
    ...ids.map(id => `- The lawyer must resolve this before the draft is used. [ref: ${id}]`)].join("\n")
}

export const happyPath: Record<string, Reply> = { [AGENT.draft]: drafter(DRAFT), [AGENT.review]: review, [AGENT.changes]: changes, [AGENT.finalize]: draftingNotes }

// Runs that plan a second document: a covering letter for the supply note,
// whose part labels are not its headings.
export const LETTER = "# Covering letter\nForm: letter\n\n## Enclosure\n\nWe enclose the supply note for signature.\n\n## Delivery\n\nThe goods will be delivered as the supply note states.\n"
export const DRAFTED = "The supplier shall deliver the goods."
export const APPLIED = "The supplier shall deliver the goods within thirty days of the order date."
export const QUESTION = "Should late delivery carry a remedy?"

// The supply note's paragraph gets the supported delivery period; other
// documents need no correction.
export const reviewFirst: Reply = (prompt, index) => listed(prompt, DRAFTED) ? review(prompt, index) : NO_FINDINGS
export const changesFirst: Reply = (prompt, index) => listed(prompt, DRAFTED) ? changes(prompt, index) : noChanges(prompt, index)

/** Whether the prompt lists a paragraph with exactly this text. */
export function listed(prompt: string, text: string) {
  return new RegExp(`^\\[p[0-9a-f]{12}\\] ${text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`, "m").test(prompt)
}

/** The pid of the listed paragraph whose text starts with this pattern. */
export function pidIn(prompt: string, start: string) {
  return new RegExp(`^\\[(p[0-9a-f]{12})\\] ${start}`, "m").exec(prompt)![1]!
}

/** Registers deliverables as the ingest service does (registerLeanDeliverable):
 * a create needs the name free and a replace the sha256 of the file there now;
 * a source, a file holding the lawyer's accepted Word revision or a failed
 * precondition is refused (undefined), as is a replace with `owned` of a file
 * that is no longer the system's to write over (replaceableDeliverable). A
 * Word file is stored with its paragraph pointers, and the sha256 of the stored
 * bytes is recorded in the registry as the system's own write (system-docx.ts).
 * `calls` lists every attempt, with the sha256 it was to replace. */
export function registering() {
  const names: string[] = []
  const calls: [string, string | null][] = []
  const registerDeliverable: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, prior, options) => {
    calls.push([name, prior])
    const file = path.join(directory, name)
    const current = existsSync(file) ? sha256Hex(readFileSync(file)) : null
    if (current !== prior || readWorkingDocx(directory, file) || registeredSource(directory, file)) return undefined
    if (prior && options?.owned && !replaceableDeliverable(directory, name)) return undefined
    const stored = name.toLowerCase().endsWith(".docx") ? withParagraphPointers(bytes) : bytes
    writeFileSync(file, stored)
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    try { db.transaction(() => recordSystemDocx(db, directory, name, sha256Hex(stored))).immediate() } finally { db.close() }
    names.push(name)
    return name
  }
  return { names, calls, registerDeliverable }
}

function sha256Hex(bytes: Uint8Array) {
  return createHash("sha256").update(bytes).digest("hex")
}

/** Whether the registry lists `file` as a source, as ingest refuses to replace it. */
function registeredSource(directory: string, file: string) {
  const registry = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(registry)) return false
  const db = new Database(registry, { readonly: true })
  try {
    if (!db.query("SELECT name FROM sqlite_master WHERE type='table' AND name='documents'").get()) return false
    return Boolean(db.query("SELECT 1 FROM documents WHERE doc_path = ? AND role = 'source'").get(file))
  } finally {
    db.close()
  }
}

// Fixture names used by the restore and deliverable tests. `prepare` may create
// the matter database first (the ingest schema); otherwise a minimal chunks
// table is used. The caller removes the directory.
export const FIXTURE_SOURCE = SOURCE
export const FIXTURE_SOURCE_TEXT = SOURCE_TEXT

export function leanFixtureMatter(prefix: string, prepare?: (directory: string) => void) {
  const directory = prepare ? createMatterDirectory(prefix) : createMatter(prefix, SOURCE_TEXT)
  if (prepare) prepare(directory)
  initLeanState(directory, { request: "Draft a short supply note from the supply terms.", ownerSessionID: "ses_owner", model: null })
  return directory
}

/** Replies by agent; a missing agent throws, so a test proves which stages ran. */
export function scriptedLeanModel(replies: Record<string, Reply> = happyLeanReplies()) {
  const model = scriptedModel(replies)
  return { ...model, agents: () => model.prompts.map(prompt => prompt.agent) }
}

export function happyLeanReplies(): Record<string, Reply> {
  return { ...happyPath }
}

function createMatterDirectory(prefix: string) {
  const directory = mkdtempSync(path.join(tmpdir(), prefix))
  mkdirSync(path.join(directory, ".dochaus"), { recursive: true, mode: 0o700 })
  return directory
}

function createMatter(prefix: string, sourceText: string) {
  const directory = createMatterDirectory(prefix)
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(SOURCE, "1", sourceText, 0, sourceText.length)
  db.close()
  return directory
}
