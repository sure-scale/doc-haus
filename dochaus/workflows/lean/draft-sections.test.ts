import { Database } from "bun:sqlite"
import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { readParagraphs } from "./checks"
import { checkSection, runDraftStage } from "./draft-sections"
import { readLeanEvents } from "./events"
import { NOT_STATED, parsePlan, readLeanPlan } from "./plan"
import { stageInputSha } from "./runner"
import { acceptStageReply, LEAN_STAGE_AGENTS, type StageSessionPort } from "./stages"
import { initLeanState, leanFile, writeLeanText } from "./state"
import { rowsBody } from "./template-coverage"

// Real plan parser, section checks, assembly, renderer and pointer writer.
// Only the draft agent's replies are scripted.

const SOURCE = "terms.docx"
const directories: string[] = []

afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "lean-sections-"))
  directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"), { recursive: true })
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(SOURCE, "1", "Notice must be given in writing.", 0, 32)
  db.close()
  return { directory, state: initLeanState(directory, { request: "Draft a short notice letter.", ownerSessionID: "ses_owner", model: null }) }
}

function scripted(reply: (text: string, index: number) => string) {
  const prompts: { sessionID: string; text: string }[] = []
  const sessions: StageSessionPort = {
    createSession: async () => "ses_draft",
    prompt: async input => {
      // The outline's coverage job finds no lawyer instruction to account for.
      if (input.agent === LEAN_STAGE_AGENTS.coverage) return JSON.stringify({ acts: [] })
      prompts.push({ sessionID: input.sessionID, text: input.text })
      expect(input.agent).toBe(LEAN_STAGE_AGENTS.draft)
      return reply(input.text, prompts.length - 1)
    },
  }
  return { sessions, prompts }
}

const OUTLINE = "Acting for: not stated\n# Notice letter\n## Background\nWhat happened.\n## Notice\nThe notice itself.\n## Next steps\nWhat follows."

describe("sectioned draft stage", () => {
  test("a presentation failure returns render feedback through the normal stage path", async () => {
    const { directory } = matter()
    const source = "# Letter\n\n## Opening\n\nText."
    expect((await acceptStageReply(directory, "d1.draft", source)).issues).toEqual([])
    const rejected = await acceptStageReply(directory, "d1.draft", source, { presentation: { source, parts: [{ heading: "Missing boundary", display_heading: "" }] } })
    expect(rejected.issues).toEqual([expect.stringContaining("Rendered draft does not contain all planned part boundaries in order")])
    expect(rejected.outputs).toEqual([])
  }, 60_000)

  test("explicit presentation changes headings without losing repeated paragraphs, tables or part attribution", async () => {
    const { directory, state } = matter()
    const outline = "Acting for: not stated\n# Notice letter\n## Opening instructions\nDisplay heading: none\nOpening text.\n## Terms &amp; drafting\nDisplay heading: **Terms**\nTerms text.\n## Closing instructions\nDisplay heading: none\nClosing text."
    const model = scripted((text, index) => {
      if (index === 0) return outline
      if (text.includes("part 1 of 3")) return "Opening instructions\n\nRepeated text."
      if (text.includes("part 2 of 3")) return "Repeated text.\n\n| Item | Value |\n| --- | --- |\n| Notice | Written |"
      return "Repeated text.\n\nWe await your reply."
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    const plan = readLeanPlan(directory)!.documents[0]!
    expect(plan.parts.map(part => part.display_heading)).toEqual(["", "**Terms**", ""])
    expect(plan.parts.map(part => part.scope)).toEqual(["Opening text.", "Terms text.", "Closing text."])
    const markdown = readFileSync(leanFile(directory, "d1/draft.md"), "utf8")
    expect(markdown).not.toContain("## Opening instructions")
    expect(markdown).not.toContain("## Closing instructions")
    expect(markdown).toContain("## **Terms**\n")
    const rows = readParagraphs(directory, "d1")
    expect(rows.filter(row => row.text === "Repeated text.").map(row => row.draft_part)).toEqual([0, 1, 2])
    expect(rows.find(row => row.text === "Opening instructions")?.draft_heading).toBe(false)
    expect(rows.find(row => row.text === "Terms")?.draft_heading).toBe(true)
    expect(rows.filter(row => ["Item", "Value", "Notice", "Written"].includes(row.text)).map(row => row.draft_part)).toEqual([1, 1, 1, 1])
    expect(rows.some(row => row.text === "Terms & drafting" || row.text === "Closing instructions")).toBe(false)
    expect(rows.every(row => row.sha256 === paragraphSha256(row.text))).toBe(true)
    const body = strFromU8(unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))["word/document.xml"]!)
    expect(body).toContain("<w:tbl>")
    const pointers = [...body.matchAll(/w:name="_dh_(p[0-9a-f]{12})"/g)].map(match => match[1])
    expect(pointers.toSorted()).toEqual(rows.map(row => row.pid).toSorted())
    const measured = rowsBody(plan, rows)
    expect(measured.partOf).toEqual(rows.map(row => row.draft_part ?? undefined))
    expect([...measured.headings].map(index => measured.paragraphs[index])).toEqual(["Notice letter", "Terms"])
    const repeated = rows.filter(row => row.text === "Repeated text.")
    const final = rows.flatMap(row => row.pid === repeated[1]!.pid ? [{ ...row, text: "Changed terms." }] : [row])
    expect(rowsBody(plan, rows, final).partOf).toEqual(measured.partOf)
    expect(model.prompts).toHaveLength(4)
    expect(model.prompts[1]!.text).toContain("the renderer omits it")
  }, 60_000)

  test("each section gets current decisions and dependency scopes without repeating the evidence bundle", async () => {
    const { directory, state } = matter()
    writeLeanText(directory, "updated-notes.md", "Keep the notice factual and use the confirmed address.")
    writeLeanText(directory, "questions.json", JSON.stringify({ questions: [{ id: "q-address", status: "answered", question: "Which address?", answer: "The address in the signed terms." }] }))
    writeLeanText(directory, "spine.json", JSON.stringify({ entries: [{ detail: "Source-backed notice requirement" }], limitations: [] }))
    const model = scripted((text, index) => {
      if (index === 0) return OUTLINE
      return text.includes("part 1 of 3") ? "The parties signed the terms." : text.includes("part 2 of 3") ? "Notice is given in writing." : "We await your reply."
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    for (const prompt of model.prompts.slice(1)) {
      expect(prompt.text).toContain("Keep the notice factual and use the confirmed address.")
      expect(prompt.text).toContain("The address in the signed terms.")
      expect(prompt.text).toContain(".dochaus/lean/spine.json")
      expect(prompt.text).toContain("What happened.")
      expect(prompt.text).toContain("What follows.")
      expect(prompt.text).not.toContain("Source-backed notice requirement")
    }
    expect(new Set(model.prompts.map(prompt => prompt.sessionID)).size).toBe(1)
  }, 60_000)

  test("drafts the outline, then each part in one session, and assembles them with paragraph pointers in the body", async () => {
    const { directory, state } = matter()
    const model = scripted((text, index) => {
      if (index === 0) return OUTLINE
      if (text.includes("part 1 of 3")) return "The parties signed the terms.\n\nThey remain in force."
      // The second part first strays into the third; the retry is clean.
      if (text.includes("part 2 of 3")) return "## Notice\n\nNotice must be given in writing.\n\n## Next steps\n\nWe will wait."
      if (text.startsWith("Your reply did not pass the draft checks:")) return "## Notice\n\n\"Notice must be given in writing\" [source: terms.docx]."
      return "## Next steps\n\n### Timing\n\nWe will wait for a reply. [open item: reply deadline]"
    })
    const progress: unknown[] = []
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, patch => progress.push(patch), "d1", () => {})

    expect(result.status).toBe("done")
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title, row.parts.length])).toEqual([["d1", "Notice letter", 3]])
    expect(model.prompts.map(prompt => prompt.sessionID)).toEqual(Array(5).fill("ses_draft"))
    expect(model.prompts[0]!.text).toContain("Reply now with the outline only")
    expect(model.prompts[3]!.text).toContain("The reply also contains other parts (\"Next steps\")")
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe([
      "# Notice letter", "",
      "## Background", "", "The parties signed the terms.", "", "They remain in force.", "",
      "## Notice", "", "\"Notice must be given in writing\" [source: terms.docx].", "",
      "## Next steps", "", "### Timing", "", "We will wait for a reply. [open item: reply deadline]", "",
    ].join("\n"))
    // Six calls, the outline's coverage job included; the hardest step (the
    // retried part) took two tries. The accepted draft's key is recorded with it.
    expect(progress.slice(-2)).toEqual([{ attempts: 2, calls: 6, feedback: [] }, { input_sha: stageInputSha(directory, "d1.draft") }])

    // Every listed paragraph carries its hidden pointer bookmark in the body docx.
    const rows = readParagraphs(directory, "d1")
    const body = strFromU8(unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))["word/document.xml"]!)
    const marks = [...body.matchAll(/w:name="_dh_(p[0-9a-f]{12})"/g)].map(match => match[1])
    expect(rows.map(row => row.text)).toContain("We will wait for a reply. [open item: reply deadline]")
    expect(marks.toSorted()).toEqual(rows.map(row => row.pid).toSorted())

    const events = readLeanEvents(directory).map(event => `${event.event}:${event.step}`)
    expect(events).toEqual([
      "model_call:outline", "checker:outline", "model_call:coverage", "checker:coverage", "coverage:coverage",
      "model_call:part 1/3", "checker:part 1/3",
      "model_call:part 2/3", "checker:part 2/3", "retry:part 2/3", "checker:part 2/3",
      "model_call:part 3/3", "checker:part 3/3",
    ])
  }, 60_000)

  test("a part that keeps failing its check stops the stage with that part's issues", async () => {
    const { directory, state } = matter()
    const model = scripted((_, index) => index === 0 ? OUTLINE : "# Notice letter\n\nEverything at once.")
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(model.prompts).toHaveLength(5)
    expect(result).toEqual({
      status: "failed",
      reason: expect.stringMatching(/^d1\.draft: part 1\/3 "Background": checker failed after 3 retries: Begin the part with the line "## Background"\./),
    })
  }, 60_000)

  test("the outline needs a title and distinct parts for each document", () => {
    expect(parsePlan("## One\n## Two").issues.join(" ")).toContain("has no \"# \" title line")
    expect(parsePlan("# Title\nNo parts here.").issues.join(" ")).toContain("\"Title\" lists no parts")
    expect(parsePlan("# Title\n## Terms\n## terms").issues.join(" ")).toContain("appear more than once")
    expect(parsePlan("# Title\n## One\n# title\n## Two").issues.join(" ")).toContain("document titles appear more than once")
    expect(parsePlan("Acting for: not stated\n# Title ##\n## One\n- covers the start\n## Two").value?.documents)
      .toEqual([{ title: "Title", purpose: "", parts: [{ heading: "One", scope: "covers the start" }, { heading: "Two", scope: "" }] }])
    // Several documents: each "# " starts one, its plain lines before the first
    // part are its purpose, and the same part heading may recur across documents.
    expect(parsePlan("Acting for: not stated\n# First\nWhat the first is for.\n## Terms\n# Second\n## Terms\nCovers terms.").value?.documents).toEqual([
      { title: "First", purpose: "What the first is for.", parts: [{ heading: "Terms", scope: "" }] },
      { title: "Second", purpose: "", parts: [{ heading: "Terms", scope: "Covers terms." }] },
    ])
  })

  test("the outline records whom the lawyer acts for from one line before the first title", () => {
    const stated = parsePlan("Acting for: the buyer\n# Title\n## One")
    expect(stated).toEqual({ issues: [], value: { documents: [expect.objectContaining({ title: "Title" })], acting_for: { party: "the buyer", stated: true } } })
    expect(parsePlan("- **Acting for:** Not stated.\n# Title\n## One")).toEqual({ issues: [], value: expect.objectContaining({ acting_for: NOT_STATED }) })
    // Without the line the outline is retried; its value stands as not stated if the last attempt still lacks it.
    for (const reply of ["# Title\n## One", "# Title\nActing for: the buyer\n## One"]) {
      const missing = parsePlan(reply)
      expect(missing.issues).toEqual([expect.stringContaining('The outline has no "Acting for:" line.')])
      expect(missing.value?.acting_for).toEqual(NOT_STATED)
    }
    // A structural defect still yields no outline.
    expect(parsePlan("Acting for: the buyer\n## One").value).toBeUndefined()
  })

  test("a missing acting-for line is retried, the outline is saved with it, and the first part is told whom the lawyer acts for", async () => {
    const { directory, state } = matter()
    writeLeanText(directory, "request.md", "Draft a notice of the terms. We act for the sender.")
    const model = scripted((text, index) => {
      if (index === 0) return OUTLINE.replace("Acting for: not stated\n", "")
      if (index === 1) return OUTLINE.replace("not stated", "the sender")
      return text.includes("part 1 of 3") ? "The parties signed the terms." : text.includes("part 2 of 3") ? "Notice is given." : "We will wait."
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    expect(model.prompts[1]!.text).toContain('The outline has no "Acting for:" line.')
    expect(readLeanPlan(directory)?.acting_for).toEqual({ party: "the sender", stated: true })
    expect(model.prompts[2]!.text).toContain("<acting_for>\nThe lawyer acts for: the sender (as the request states).\n</acting_for>")
    expect(model.prompts[3]!.text).not.toContain("<acting_for>")
  }, 60_000)

  test("an outline that never states the line is accepted on the last attempt as not stated", async () => {
    const { directory, state } = matter()
    const model = scripted((text, index) => index < 4 ? OUTLINE.replace("Acting for: not stated\n", "")
      : text.includes("part 1 of 3") ? "The parties signed the terms." : text.includes("part 2 of 3") ? "Notice is given." : "We will wait.")
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    expect(readLeanPlan(directory)?.acting_for).toEqual(NOT_STATED)
    expect(model.prompts[4]!.text).toContain("The request does not say whom the lawyer acts for.")
  }, 60_000)

  test("a part reply is checked against its outlined heading", () => {
    const outline = { title: "T", parts: [{ heading: "One", scope: "" }, { heading: "Two", scope: "" }] }
    expect(checkSection("Text only.", outline, 0)).toEqual({ issues: [], value: "## One\n\nText only." })
    expect(checkSection("## One\n\nText.\n\n## **Two**\n\nMore text.", outline, 0).issues).toEqual([expect.stringContaining('use "### " or lower headings inside it')])
    expect(checkSection("## One\n\nText.\n\n## T&#119;o\n\nMore text.", outline, 0).issues).toEqual([expect.stringContaining('use "### " or lower headings inside it')])
    expect(checkSection("## One\n\nText.\n\n## One\n\nMore text.", outline, 0).issues).toEqual([expect.stringContaining('use "### " or lower headings inside it')])
    expect(checkSection("## One\n\n### Sub\n", outline, 0).issues).toEqual(["The part \"One\" has no text under its heading."])
    expect(checkSection("## One\n\nText.\n\n# T", outline, 0).issues[0]).toContain("\"# \" title line")
    expect(checkSection("", outline, 1).issues[0]).toContain("The reply was empty")
  })
})
