import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { LEAN_ACTS_SCHEMA } from "./act-record"
import { runApply } from "./apply-stage"
import { checkChanges } from "./change-checks"
import { readParagraphs } from "./checks"
import { runDraftStage } from "./draft-sections"
import { finalGate } from "./gate"
import { parsePlan, writeLeanPlan } from "./plan"
import { READER_CHECK_PERMISSIONS, parseReaderHits, readReaderChecks, readerCheckKey, readerCheckPrompt, readerItems, saveReaderCheck } from "./reader-check"
import { changesRepairSections } from "./repairs"
import { stageInputSha } from "./runner"
import { acceptStageReply, type PermissionRule, type StageSessionPort } from "./stages"
import { readLeanState, replaceStageOpenItems, updateLeanStage, writeLeanText } from "./state"
import { AGENT, cleanupMatters, labelledPart, matter } from "./test-support"

// A document's text speaks to its reader; a sentence written for the lawyer is
// a hit the change cycle repairs. Synthetic text only.

const workspace = process.env.WORKSPACE_ROOT
const roots: string[] = []
afterEach(() => {
  cleanupMatters()
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
  if (workspace === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = workspace
})

const NOTE = "Please check the reply date before this letter goes out."
const OUTLINE = "Acting for: not stated\n# Notice letter\nForm: letter\n## Background\nWhat happened.\n## Next steps\nWhat follows."

function rows(texts: string[]) {
  return texts.map((text, index) => ({ pid: `p${String(index + 1).padStart(12, "0")}`, text, sha256: paragraphSha256(text) }))
}

const hit = (pid: string, sentence: string) => ({ pid, sentence, addressed_to: "the lawyer", keep: "that a reply is awaited" })

function letter(directory: string) {
  const plan = parsePlan(OUTLINE)
  expect(plan.issues).toEqual([])
  writeLeanPlan(directory, plan.value!.documents)
}

function firmLegend(legend: string) {
  const root = mkdtempSync(path.join(tmpdir(), "reader-check-prefs-"))
  roots.push(root)
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences", "preferences.json"), JSON.stringify({ adviceLegend: legend }))
  process.env.WORKSPACE_ROOT = root
}

/** A session port scripted by agent, recording each session and prompt. */
function port(replies: Record<string, (text: string, index: number) => string>) {
  const sessions: { id: string; agent: string; permissions?: PermissionRule[] }[] = []
  const prompts: { agent: string; sessionID: string; text: string }[] = []
  const sessionPort: StageSessionPort = {
    async createSession(input) {
      const id = `ses_${sessions.length + 1}`
      sessions.push({ id, agent: input.agent, permissions: input.permissions })
      return id
    },
    async prompt(input) {
      const index = prompts.filter(row => row.agent === input.agent).length
      prompts.push({ agent: input.agent, sessionID: input.sessionID, text: input.text })
      const reply = replies[input.agent]
      if (!reply) throw new Error(`No scripted reply for ${input.agent}`)
      return reply(input.text, index)
    },
  }
  return { port: sessionPort, sessions, prompts }
}

/** A reader job's reply: the paragraph of its part that holds NOTE, as a hit. */
function readerReply(text: string) {
  const part = /<part>\n([\s\S]*?)\n<\/part>/.exec(text)![1]!
  const row = part.split("\n\n").find(line => line.includes(NOTE))
  const pid = row && /^\[(p[0-9a-f]+)\]/.exec(row)?.[1]
  return JSON.stringify({ hits: pid ? [hit(pid, NOTE)] : [] })
}

describe("parseReaderHits", () => {
  const part = rows(["We will deliver the goods on the date agreed. " + NOTE, "Payment follows delivery."])
  const reply = (hits: unknown[]) => JSON.stringify({ hits })

  test("a hit's paragraph must be in the part and its sentence in that paragraph word for word; no hit is a valid reply", () => {
    expect(parseReaderHits(reply([hit(part[0]!.pid, NOTE)]), part)).toEqual({ issues: [], value: [hit(part[0]!.pid, NOTE)] })
    // Spacing and quote marks are read as the paragraph shows them.
    expect(parseReaderHits(reply([hit(part[0]!.pid, "Please  check the reply date before this letter goes out.")]), part).value).toHaveLength(1)
    expect(parseReaderHits(reply([]), part)).toEqual({ issues: [], value: [] })
    expect(parseReaderHits(reply([hit("p000000000099", NOTE)]), part).issues).toEqual([expect.stringContaining('hits[0]: pid "p000000000099" is not a paragraph of this part')])
    expect(parseReaderHits(reply([hit(part[1]!.pid, NOTE)]), part).issues).toEqual([expect.stringContaining(`hits[0] ([${part[1]!.pid}]): the sentence is not in that paragraph word for word`)])
    expect(parseReaderHits(reply([{ pid: part[0]!.pid, sentence: NOTE }]), part).issues).toEqual([expect.stringContaining("hits[0]: Reply with the JSON object only")])
    expect(parseReaderHits("Not JSON", part).issues).toEqual([expect.stringContaining("Reply with the JSON object only")])
    expect(parseReaderHits(JSON.stringify({ hits: "none" }), part).issues).toEqual([expect.stringContaining("Reply with the JSON object only")])
  })

  test("on the last attempt the verified hits stand and the rest are dropped", () => {
    const last = parseReaderHits(reply([hit(part[0]!.pid, NOTE), hit(part[1]!.pid, NOTE), hit(part[0]!.pid, NOTE)]), part, [], { final: true })
    expect(last.value).toEqual([hit(part[0]!.pid, NOTE)])
    expect(last.issues).toHaveLength(1)
    expect(parseReaderHits(reply([hit(part[0]!.pid, NOTE), hit(part[1]!.pid, NOTE)]), part).value).toBeUndefined()
  })

  test("a sentence inside a marker, a placeholder or quoted words, or taken from the furniture or the original, is no hit", () => {
    const texts = [
      "The fee is 100. [open item: confirm the fee with the client before sending]",
      "Notice goes to [insert the notice address before sending].",
      "The supplier wrote: “Please verify this figure before sending.”",
      "Please verify the dates before sending. [open item: the dates]",
      "Privileged and confidential: prepared for the client's review.",
    ]
    const part = rows(texts)
    const kept = ["Privileged and confidential: prepared for the client's review."]
    const checked = parseReaderHits(reply([
      hit(part[0]!.pid, "confirm the fee with the client before sending"),
      hit(part[1]!.pid, "insert the notice address before sending"),
      hit(part[2]!.pid, "Please verify this figure before sending."),
      hit(part[3]!.pid, "Please verify the dates before sending."),
      hit(part[4]!.pid, "Privileged and confidential: prepared for the client's review."),
    ]), part, kept)
    expect(checked).toEqual({ issues: [], value: [hit(part[3]!.pid, "Please verify the dates before sending.")] })
  })

  test("a mixed prose-and-marker hit must quote the operative words alone, without hiding a real instruction beside the marker", () => {
    const recipient = "You may send your reply to the address above."
    const marker = "[open item: confirm the reply address before sending]"
    const part = rows([`${recipient} ${marker} ${NOTE}`])
    const mixed = hit(part[0]!.pid, `${recipient} ${marker}`)
    const real = hit(part[0]!.pid, NOTE)
    // Merely sharing a paragraph with a marker never shelters actual lawyer-facing prose.
    expect(parseReaderHits(reply([real]), part)).toEqual({ issues: [], value: [real] })
    const checked = parseReaderHits(reply([mixed, real]), part, [], { final: true })
    expect(checked.issues).toEqual([expect.stringContaining("crosses a lawyer-instruction marker")])
    expect(checked.value).toEqual([real])
    expect(parseReaderHits(reply([mixed]), part).value).toBeUndefined()
  })

  test("a marker in the middle of prose cannot manufacture a word-for-word hit by joining its two sides", () => {
    const part = rows(["Please verify [open item: identify the source] the reply date before sending."])
    expect(parseReaderHits(reply([hit(part[0]!.pid, "Please verify the reply date before sending.")]), part).issues)
      .toEqual([expect.stringContaining("the sentence is not in that paragraph word for word")])
    expect(parseReaderHits(reply([hit(part[0]!.pid, "the reply date before sending.")]), part).value)
      .toEqual([hit(part[0]!.pid, "the reply date before sending.")])
  })
})

test("the reader prompt separates exact operative passages from all supported lawyer markers, retaining the canonical paragraph", () => {
  const plan = { ...parsePlan(OUTLINE).value!.documents[0]!, id: "d1" as const }
  const part = rows([
    "You may reply [insert reply address] by the date stated. [open item: confirm the date]",
    `${NOTE} [ATTORNEY NOTE: review the reply deadline]`,
  ])
  const prompt = readerCheckPrompt(plan, "", part, [])
  expect(prompt).toContain(`<part>\n[${part[0]!.pid}] ${part[0]!.text}`)
  const operative = JSON.parse(/<operative_prose>\n([\s\S]*?)\n<\/operative_prose>/.exec(prompt)![1]!)
  const instructions = JSON.parse(/<lawyer_instructions>\n([\s\S]*?)\n<\/lawyer_instructions>/.exec(prompt)![1]!)
  expect(operative).toEqual([
    { pid: part[0]!.pid, passages: ["You may reply", "by the date stated."] },
    { pid: part[1]!.pid, passages: [NOTE] },
  ])
  expect(instructions).toEqual([
    { pid: part[0]!.pid, passages: ["[insert reply address]", "[open item: confirm the date]"] },
    { pid: part[1]!.pid, passages: ["[ATTORNEY NOTE: review the reply deadline]"] },
  ])
  expect(prompt).toContain("does not resolve its factual gap")
  expect(prompt).toContain("Do not join passages across a marker")
})

describe("reader items", () => {
  test("a draft's hit reaches the change cycle as a required repair, in plain words", () => {
    const directory = matter()
    letter(directory)
    const draft = rows(["We will wait for a reply. " + NOTE])
    writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(draft))
    const items = readerItems("draft", "d1", [hit(draft[0]!.pid, NOTE)])
    expect(items).toEqual([expect.objectContaining({ kind: "reader", stage: "draft", ref: { document: "d1", pid: draft[0]!.pid, statement: NOTE } })])
    expect(items[0]!.detail).toContain(NOTE)
    expect(items[0]!.detail).not.toContain(draft[0]!.pid)
    replaceStageOpenItems(directory, "draft", "d1", items)
    const sections = changesRepairSections(directory, "d1").join("\n")
    expect(sections).toContain(`${items[0]!.id} [${draft[0]!.pid}]`)
    expect(sections).toContain("written for the lawyer rather than the document's reader")
  })

  test("a hit left in the final body holds delivery until the drafting notes name it; one no longer there is resolved", () => {
    const directory = matter()
    letter(directory)
    const final = rows(["We will wait for a reply. " + NOTE])
    writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(final))
    const items = readerItems("apply", "d1", [hit(final[0]!.pid, NOTE)])
    replaceStageOpenItems(directory, "apply", "d1", items)
    writeLeanText(directory, "drafting-notes.md", "# Drafting notes\n\nNothing to report.\n")
    expect(finalGate(directory).blocking.map(item => item.kind)).toContain("reader")
    writeLeanText(directory, "drafting-notes.md", `# Drafting notes\n\nThe letter still asks you to check the reply date [ref: ${items[0]!.id}].\n`)
    expect(finalGate(directory).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
    writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows(["We will wait for a reply."])))
    writeLeanText(directory, "drafting-notes.md", "# Drafting notes\n\nNothing to report.\n")
    expect(finalGate(directory).items).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "resolved" }))
  })
})

describe("the draft step's reader check", () => {
  test("launches one job per drafted part in a session of its own, keeps its hits as reader items, and does not ask a settled part again", async () => {
    const directory = matter()
    writeLeanText(directory, "acts.json", JSON.stringify({ schema: LEAN_ACTS_SCHEMA, status: "none", acts: [], verify: null, deliverables: [], issues: [] }))
    const model = port({
      [AGENT.draft]: (text, index) => index === 0 ? OUTLINE : text.includes("part 1 of 2") ? labelledPart(text, "The parties signed the terms.") : labelledPart(text, `We will wait for a reply. ${NOTE}`),
      [AGENT.reader]: readerReply,
    })
    const state = readLeanState(directory)!
    const result = await runDraftStage(directory, state, { sessions: model.port }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    if (result.status !== "done") throw new Error(result.reason)
    const readers = model.prompts.filter(row => row.agent === AGENT.reader)
    expect(readers).toHaveLength(2)
    expect(new Set(readers.map(row => row.sessionID)).size).toBe(2)
    expect(model.sessions.filter(row => row.agent === AGENT.reader).map(row => row.permissions)).toEqual([READER_CHECK_PERMISSIONS, READER_CHECK_PERMISSIONS])
    const [background, next] = readers.map(row => row.text).toSorted((a, b) => (a.includes(NOTE) ? 1 : 0) - (b.includes(NOTE) ? 1 : 0))
    for (const text of [background!, next!]) {
      expect(text).toContain("<document>\n")
      expect(text).toContain('"reader": "the person it is addressed to"')
      expect(text).toContain("<part>\n[p")
      expect(text).toContain("<question>\n")
      expect(text).toContain("<output_contract>\n")
    }
    expect(background).toContain("The parties signed the terms.")
    expect(background).not.toContain(NOTE)
    expect(next).not.toContain("The parties signed the terms.")
    const row = readParagraphs(directory, "d1").find(row => row.text.includes(NOTE))!
    const items = result.openItems.filter(item => item.kind === "reader")
    expect(items).toEqual([expect.objectContaining({ stage: "draft", ref: { document: "d1", pid: row.pid, statement: NOTE } })])
    expect(Object.values(readReaderChecks(directory, "d1"))).toHaveLength(2)

    // The accepted draft is kept on a re-run from its key, and so are its settled parts.
    const kept = updateLeanStage(directory, readLeanState(directory)!, "d1.draft", { status: "done", outputs: result.outputs, input_sha: stageInputSha(directory, "d1.draft") })
    const again = await runDraftStage(directory, kept, { sessions: model.port }, () => {}, "d1", () => {})
    expect(again.status).toBe("done")
    expect(model.prompts.filter(row => row.agent === AGENT.reader)).toHaveLength(2)
    expect(model.prompts.filter(row => row.agent === AGENT.draft)).toHaveLength(3)
    expect(again.status === "done" ? again.openItems.filter(item => item.kind === "reader") : []).toEqual(items)
  }, 60_000)

  test("the firm's legend for the form is no hit, and the job is told it prints on every page", async () => {
    const directory = matter()
    firmLegend("Privileged and confidential")
    writeLeanText(directory, "acts.json", JSON.stringify({ schema: LEAN_ACTS_SCHEMA, status: "none", acts: [], verify: null, deliverables: [], issues: [] }))
    const model = port({
      [AGENT.draft]: (text, index) => index === 0 ? "Acting for: not stated\n# Advice note\nForm: advice\n## Advice\nThe advice." : labelledPart(text, "Privileged and confidential. The notice was valid."),
      [AGENT.reader]: text => {
        const pid = /<part>\n\[(p[0-9a-f]+)\] Privileged and confidential\./.exec(text)?.[1]
        return JSON.stringify({ hits: pid ? [hit(pid, "Privileged and confidential.")] : [] })
      },
    })
    const result = await runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    const reader = model.prompts.find(row => row.agent === AGENT.reader)!
    expect(reader.text).toContain("<furniture>\nLines code prints on every page from the firm's settings:\nPrivileged and confidential\n</furniture>")
    expect(result.status === "done" ? result.openItems.filter(item => item.kind === "reader") : []).toEqual([])
  }, 60_000)
})

describe("the apply step's reader check", () => {
  test("refreshes current parts after a contract change, retains genuine standing defects and then reuses unchanged parts", async () => {
    const directory = matter()
    letter(directory)
    const recipient = "We will wait for a reply. [open item: confirm the reply date]"
    expect((await acceptStageReply(directory, "d1.draft", `# Notice letter\n\n## Background\n\nThe parties signed the terms.\n\n## Next steps\n\n${recipient} ${NOTE}\n`)).issues).toEqual([])
    const draft = readParagraphs(directory, "d1")
    const background = draft.find(row => row.text === "The parties signed the terms.")!
    const next = draft.find(row => row.text.includes(NOTE))!
    const obsolete = hit(next.pid, recipient)
    const found = readerItems("draft", "d1", [obsolete, hit(next.pid, NOTE)])
    replaceStageOpenItems(directory, "draft", "d1", found)
    saveReaderCheck(directory, "d1", readerCheckKey("Previous reader contract"), [obsolete, hit(next.pid, NOTE)])
    const amend = (changes: [string, string][]) => {
      const checked = checkChanges(directory, JSON.stringify({ amendments: changes.map(([pid, updated]) => ({ pid, updated, reason: "Review finding.", sources: [] })), questions: [] }), readParagraphs(directory, "d1"), { document: "d1" })
      expect(checked.issues).toEqual([])
      writeLeanText(directory, "d1/changes.json", `${JSON.stringify(checked.changes, null, 2)}\n`)
    }
    // Each change rests on drafting judgment, so it carries a new marker for the lawyer.
    const changedBackground = "The parties signed the terms on the date stated. [open item: confirm the signing date]"
    amend([[background.pid, changedBackground]])
    const model = port({ [AGENT.reader]: readerReply })
    const applied = await runApply(directory, readLeanState(directory)!, "d1", { sessions: model.port })
    expect(applied.status).toBe("done")
    if (applied.status !== "done") throw new Error(applied.reason)
    expect(model.prompts.map(row => row.agent)).toEqual([AGENT.reader, AGENT.reader])
    expect(model.prompts[0]!.text).toContain("on the date stated")
    expect(model.prompts[0]!.text).not.toContain(NOTE)
    expect(model.prompts[1]!.text).toContain(recipient)
    expect(applied.openItems.filter(item => item.kind === "reader")).toEqual([expect.objectContaining({ stage: "apply", detail: found[1]!.detail, ref: { document: "d1", pid: next.pid, statement: NOTE } })])

    // A change that rewrites the sentence's own part is checked again there; the draft's hit is not carried.
    amend([[background.pid, changedBackground], [next.pid, "We will wait for a reply. [open item: confirm the reply date and address]"]])
    const again = await runApply(directory, readLeanState(directory)!, "d1", { sessions: model.port })
    expect(again.status).toBe("done")
    if (again.status !== "done") throw new Error(again.reason)
    expect(model.prompts.map(row => row.agent)).toEqual([AGENT.reader, AGENT.reader, AGENT.reader])
    expect(model.prompts[2]!.text).toContain("We will wait for a reply.")
    expect(model.prompts[2]!.text).not.toContain("The parties signed the terms")
    expect(again.openItems.filter(item => item.kind === "reader")).toEqual([])

    // Without a session port nothing is asked; a standing hit is still carried.
    amend([[background.pid, changedBackground]])
    const direct = await runApply(directory, readLeanState(directory)!, "d1")
    expect(direct.status).toBe("done")
    expect(model.prompts).toHaveLength(3)
    expect(direct.status === "done" ? direct.openItems.filter(item => item.kind === "reader") : []).toHaveLength(2)
  }, 60_000)
})
