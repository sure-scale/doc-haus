import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
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

export function scriptedModel(replies: Record<string, Reply>) {
  const prompts: { agent: string; sessionID: string; text: string }[] = []
  const sessions: Parameters<StageSessionPort["createSession"]>[0][] = []
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
        // The outline's coverage job finds no lawyer instruction to account for.
        [AGENT.coverage]: () => JSON.stringify({ acts: [] }),
        [AGENT["change-plan"]]: () => JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [] }),
        [AGENT["change-prepare"]]: () => JSON.stringify({ amendments: [], questions: [] }),
        [AGENT["change-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [] }),
        [AGENT.discrepancy]: () => JSON.stringify({ entries: [], limitations: [] }),
        [AGENT.spine]: () => JSON.stringify({ entries: [], limitations: [] }),
        [AGENT["review-independent"]]: () => "No changes needed.",
        [AGENT["review-wording"]]: () => "No changes needed.",
        [AGENT["review-client-disadvantage"]]: () => "No changes needed.",
        [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [] }),
        [AGENT.intake]: () => JSON.stringify({ ready: true, jurisdiction: "No external legal determination needed for this supplied-source fixture", scope: "Reproduce supplied terms only" }),
        [AGENT.research]: () => JSON.stringify({ status: "not-needed", findings: [], limitations: ["This fixture requests a note from supplied terms only."] }),
        [AGENT["research-review"]]: () => JSON.stringify({ approved: [], rejected: [], limitations: [] }),
      }
      const reply = replies[input.agent] ?? preparation[input.agent]
      if (!reply) throw new Error(`No scripted reply for ${input.agent}.`)
      return reply(input.text, prompts.filter(prompt => prompt.agent === input.agent).length - 1, input.sessionID)
    },
  }
  const count = (agent: string) => prompts.filter(prompt => prompt.agent === agent).length
  return { port, prompts, sessions, count, counts: () => [AGENT.draft, AGENT.review, AGENT.changes, AGENT.finalize].map(count) }
}

/** The draft agent for whole Markdown documents, in drafting order: their "# "
 * titles and "## " parts answer the outline prompt (and any retry of it), after
 * an "Acting for: not stated" line, then each part prompt gets the text of the
 * part it names in the document it names. Each assembled draft is the document
 * itself. */
export function drafter(...documents: string[]): Reply {
  const planned = documents.map(document => ({
    title: /^# (.+)$/m.exec(document)?.[1],
    parts: document.split(/^(?=## )/m).slice(1).map(part => part.trim()),
  }))
  if (!planned.length || planned.some(row => !row.title || !row.parts.length))
    throw new Error('A scripted draft needs a "# " title and at least one "## " part.')
  return prompt => {
    const named = /^Draft part \d+ of \d+ of "([^"]*)": (.+)$/m.exec(prompt)
    if (!named) return ["Acting for: not stated", ...planned.flatMap(row => [`# ${row.title}`, ...row.parts.map(part => part.split("\n")[0]!)])].join("\n")
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

export const review: Reply = prompt =>
  `- [${pidOf(prompt, "The supplier shall deliver the goods.")}] The delivery period is missing. The terms say "within thirty days of the order date" [source: ${SOURCE}].`

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

// Names every open item id the coordinator lists, in every document's block, as
// the drafting-notes contract asks.
export const draftingNotes: Reply = prompt => {
  const blocks = [...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].map(match => match[1]!)
  const ids = blocks.flatMap(block => [...block.matchAll(/^- (\S+):/gm)].map(match => match[1]))
  return ["# Drafting notes", "", "The draft states the delivery period from the supply terms.", "",
    ...ids.map(id => `- ${id}: the lawyer must resolve this before the draft is used.`)].join("\n")
}

export const happyPath: Record<string, Reply> = { [AGENT.draft]: drafter(DRAFT), [AGENT.review]: review, [AGENT.changes]: changes, [AGENT.finalize]: draftingNotes }

// Runs that plan a second document: a covering letter for the supply note.
export const LETTER = "# Covering letter\n\n## Enclosure\n\nWe enclose the supply note for signature.\n\n## Delivery\n\nThe goods will be delivered as the supply note states.\n"
export const DRAFTED = "The supplier shall deliver the goods."
export const APPLIED = "The supplier shall deliver the goods within thirty days of the order date."
export const QUESTION = "Should late delivery carry a remedy?"

// The supply note's paragraph gets the supported delivery period; other
// documents need no correction.
export const reviewFirst: Reply = (prompt, index) => listed(prompt, DRAFTED) ? review(prompt, index) : "No corrections are needed."
export const changesFirst: Reply = (prompt, index) => listed(prompt, DRAFTED) ? changes(prompt, index) : noChanges(prompt, index)

/** Whether the prompt lists a paragraph with exactly this text. */
export function listed(prompt: string, text: string) {
  return new RegExp(`^\\[p[0-9a-f]{12}\\] ${text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`, "m").test(prompt)
}

/** The pid of the listed paragraph whose text starts with this pattern. */
export function pidIn(prompt: string, start: string) {
  return new RegExp(`^\\[(p[0-9a-f]{12})\\] ${start}`, "m").exec(prompt)![1]!
}

/** Registers deliverables by writing them into the matter, as the ingest service does. */
export function registering() {
  const names: string[] = []
  const registerDeliverable: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes) => {
    writeFileSync(path.join(directory, name), bytes)
    names.push(name)
    return name
  }
  return { names, registerDeliverable }
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
