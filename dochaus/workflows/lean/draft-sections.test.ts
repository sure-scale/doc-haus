import { Database } from "bun:sqlite"
import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { renderMarkdownToDocx } from "../../lib/render"
import { readParagraphs, shownText } from "./checks"
import { assembleDraft, checkSection, runDraftStage, sectionPrompt } from "./draft-sections"
import { readLeanEvents } from "./events"
import { NOT_STATED, leanPlanIssue, operativeDocument, parsePlan, readLeanPlan, writeLeanPlan, type PlannedDocument } from "./plan"
import { stagePrompt } from "./prompts"
import { stageInputSha } from "./runner"
import { acceptStageReply, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, outlineBasis, type StageSessionPort } from "./stages"
import { initLeanState, leanFile, updateLeanStage, writeLeanState, writeLeanText, type LeanState } from "./state"
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

/** Runs `run` with these firm settings saved in a fresh workspace, dates written day first. */
async function withFirm<T>(preferences: Record<string, unknown>, run: () => Promise<T>): Promise<T> {
  const workspace = mkdtempSync(path.join(tmpdir(), "lean-furniture-prefs-"))
  directories.push(workspace)
  mkdirSync(path.join(workspace, ".preferences"))
  writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify({ dateFormat: "day-month-year", ...preferences }))
  const root = process.env.WORKSPACE_ROOT
  process.env.WORKSPACE_ROOT = workspace
  try {
    return await run()
  } finally {
    if (root === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = root
  }
}

/** The run started on 20 October 2026, the date its draft line reads. */
const dated = (directory: string, state: LeanState) => writeLeanState(directory, { ...state, created_at: "2026-10-20T12:00:00.000Z" })

/** A rendered body's header and footer lines by reference, keyed "header:default" and so on. */
function furnitureLines(parts: Record<string, Uint8Array>) {
  const relationships = strFromU8(parts["word/_rels/document.xml.rels"]!)
  const target = (id: string) => new RegExp(`<Relationship [^>]*Id="${id}"[^>]*Target="([^"]+)"`).exec(relationships)?.[1]
  const lines = (xml: string) => [...xml.matchAll(/<w:p>([\s\S]*?)<\/w:p>/g)].map(([, paragraph]) => [...paragraph!.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(run => run[1]).join(""))
  return Object.fromEntries([...strFromU8(parts["word/document.xml"]!).matchAll(/<w:(header|footer)Reference w:type="(\w+)" r:id="(\w+)"/g)]
    .map(([, kind, type, id]) => [`${kind}:${type}`, lines(strFromU8(parts[`word/${target(id!)}`]!))]))
}

/** The rendered body's paragraphs: text, heading style and the spacing before it. */
function styledParagraphs(directory: string) {
  const body = strFromU8(unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))["word/document.xml"]!)
  return [...body.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map(([xml]) => [
    [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join(""),
    /<w:pStyle w:val="([^"]+)"/.exec(xml)?.[1], /<w:spacing w:before="(\d+)"/.exec(xml)?.[1],
  ])
}

// A notice letter whose text uses headings.
const OUTLINE = "Acting for: not stated\n# Notice letter\nForm: letter\nHeadings: shown\n## Background\nWhat happened.\n## Notice\nThe notice itself.\n## Next steps\nWhat follows."

describe("sectioned draft stage", () => {
  test("a part boundary the Word body does not show is presented without it; the draft is not failed for it", async () => {
    const { directory } = matter()
    const source = "# Letter\n\n## Opening\n\nText."
    expect((await acceptStageReply(directory, "d1.draft", source)).issues).toEqual([])
    const accepted = await acceptStageReply(directory, "d1.draft", source,
      { presentation: { source, parts: [{ heading: "Opening", display_heading: "", part: 0 }, { heading: "Missing boundary", display_heading: "", part: 1 }] } })
    expect(accepted.issues).toEqual([])
    expect(readParagraphs(directory, "d1").map(row => [row.text, row.draft_part, row.draft_heading])).toEqual([["Letter", null, true], ["Text.", 0, false]])
  }, 60_000)

  test("part labels whose rendered heading reads differently from their Markdown are still found, left out and attributed", async () => {
    const { directory, state } = matter()
    const labels = ["Fees*", "Recitals ~~old~~", "Execution page <!-- note -->"]
    const outline = `Acting for: not stated\n# Deed\nForm: other\n${labels.map(label => `## ${label}\nWhat it covers.`).join("\n")}`
    const model = scripted((text, index) => {
      if (index === 0) return outline
      const at = labels.findIndex((_, part) => text.includes(`part ${part + 1} of 3`))
      return `## ${labels[at]}\n\n### ${at + 1}. Clause\n\nText ${at + 1}.`
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    expect(readParagraphs(directory, "d1").map(row => [shownText(row), row.draft_part, row.draft_heading])).toEqual([
      ["Deed", null, true], ["1. Clause", 0, false], ["Text 1.", 0, false], ["2. Clause", 1, false], ["Text 2.", 1, false], ["3. Clause", 2, false], ["Text 3.", 2, false],
    ])
    expect(readLeanEvents(directory).filter(event => event.step === "presentation")).toEqual([])
  }, 60_000)

  test("a part label the Word body does not render as a heading leaves the label as rendered, counts its text with the part before, and is recorded", async () => {
    const { directory, state } = matter()
    const outline = "Acting for: not stated\n# Deed\nForm: other\n## Background\nWhat happened.\n## ---\nWhat follows."
    const model = scripted((text, index) => index === 0 ? outline : text.includes("part 1 of 2") ? "## Background\n\nThe parties met." : "## ---\n\nThe deed follows.")
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    const rows = readParagraphs(directory, "d1")
    expect(rows.find(row => row.text === "The deed follows.")?.draft_part).toBe(0)
    expect(rows.some(row => row.draft_part === 1)).toBe(false)
    expect(readLeanEvents(directory).filter(event => event.step === "presentation").map(event => [event.event, event.issues])).toEqual([["checker", [
      'The Word body could not place the part "---": its label stays as rendered, and its paragraphs count with the part before it when the draft is checked against a template.',
    ]]])
  }, 60_000)

  test("explicit presentation changes headings without losing repeated paragraphs, tables or part attribution", async () => {
    const { directory, state } = matter()
    // A letter's own Display heading lines stand, the visible one included.
    const outline = "Acting for: not stated\n# Notice letter\nForm: letter\n## Opening instructions\nDisplay heading: none\nOpening text.\n## Terms &amp; drafting\nDisplay heading: **Terms**\nTerms text.\n## Closing instructions\nDisplay heading: none\nClosing text."
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
    // A letter's planned title does not print.
    expect([...measured.headings].map(index => measured.paragraphs[index])).toEqual(["Terms"])
    const repeated = rows.filter(row => row.text === "Repeated text.")
    const final = rows.flatMap(row => row.pid === repeated[1]!.pid ? [{ ...row, text: "Changed terms." }] : [row])
    expect(rowsBody(plan, rows, final).partOf).toEqual(measured.partOf)
    expect(model.prompts).toHaveLength(4)
    expect(model.prompts[1]!.text).toContain("the renderer omits it")
  }, 60_000)

  test("a new outline keeps part labels internal: the drafter's own headings stand alone, the title prints once, and the text stays attributed to its parts", async () => {
    const { directory, state } = matter()
    const outline = "Acting for: not stated\n# Rent Deposit Deed\nForm: other\n## Parties and recitals\nWho gives the deed.\n## Deposit and interest\nHow the deposit is held."
    const model = scripted((text, index) => {
      if (index === 0) return outline
      if (text.includes("part 1 of 2")) return "## Parties and recitals\n\n**RENT DEPOSIT DEED**\n\n**1. Parties**\n\n1.1 The landlord and the tenant."
      if (text.startsWith("Your reply did not pass the draft checks:")) return "## Parties and recitals\n\n**1. Parties**\n\n1.1 The landlord and the tenant."
      return "## Deposit and interest\n\n**2. Deposit and interest**\n\n2.1 Text."
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    const plan = readLeanPlan(directory)!.documents[0]!
    expect(plan.parts.map(part => part.display_heading)).toEqual(["", ""])
    expect(model.prompts[1]!.text).toContain("That first line is an internal part label; the renderer omits it. Write the headings the document itself needs")
    expect(model.prompts[2]!.text).toContain('The part opens with the document title "Rent Deposit Deed".')
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe("# Rent Deposit Deed\n\n**1. Parties**\n\n1.1 The landlord and the tenant.\n\n**2. Deposit and interest**\n\n2.1 Text.\n")
    const rows = readParagraphs(directory, "d1")
    expect(rows.map(row => [row.text, row.draft_part, row.draft_heading])).toEqual([
      ["Rent Deposit Deed", null, true],
      ["1. Parties", 0, false], ["1.1 The landlord and the tenant.", 0, false],
      ["2. Deposit and interest", 1, false], ["2.1 Text.", 1, false],
    ])
    expect(rowsBody(plan, rows).partOf).toEqual([undefined, 0, 0, 1, 1])
  }, 60_000)

  test("a part whose label is not printed has its own headings one level up in the Markdown and the Word body; a part with a printed heading keeps its levels", async () => {
    const { directory, state } = matter()
    const outline = "Acting for: not stated\n# Supply Agreement\nForm: agreement\n## Parties\nWho is bound.\n## Goods\nDisplay heading: Schedule\nWhat is supplied."
    const model = scripted((text, index) => {
      if (index === 0) return outline
      if (text.includes("part 1 of 2")) return "## Parties\n\n### 1. Parties\n\n1.1 The supplier and the buyer.\n\n#### 1.2 Definitions\n\nWords have their defined meanings."
      return "## Goods\n\n### Goods supplied\n\nThe goods in the order."
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe("# Supply Agreement\n\n## 1. Parties\n\n1.1 The supplier and the buyer.\n\n" +
      "### 1.2 Definitions\n\nWords have their defined meanings.\n\n## Schedule\n\n### Goods supplied\n\nThe goods in the order.\n")
    // Text, heading style and the spacing rendering gives that style; clause
    // numbers are Word list numbering, not text.
    expect(styledParagraphs(directory)).toEqual([
      ["Supply Agreement", "Heading1", "320"],
      ["Parties", "Heading2", "280"], ["The supplier and the buyer.", undefined, undefined],
      ["Definitions", "Heading3", "220"], ["Words have their defined meanings.", undefined, undefined],
      ["Schedule", "Heading2", "280"], ["Goods supplied", "Heading3", "220"], ["The goods in the order.", undefined, undefined],
    ])
    expect(readParagraphs(directory, "d1").map(row => [shownText(row), row.draft_part, row.draft_heading])).toEqual([
      ["Supply Agreement", null, true],
      ["1. Parties", 0, false], ["1.1 The supplier and the buyer.", 0, false], ["1.2 Definitions", 0, false], ["Words have their defined meanings.", 0, false],
      ["Schedule", 1, true], ["Goods supplied", 1, false], ["The goods in the order.", 1, false],
    ])
  }, 60_000)

  test("a part whose label is not printed and whose headings start below \"### \" has its shallowest heading at the second level in the Markdown and the Word body", async () => {
    const { directory, state } = matter()
    const outline = "Acting for: not stated\n# Payment Demand\nForm: other\n## Demand\nWhat is demanded.\n## Terms\nThe terms."
    const model = scripted((text, index) => {
      if (index === 0) return outline
      if (text.includes("part 1 of 2")) return "## Demand\n\n#### Demand for payment\n\nPay the sum due.\n\n##### Interest\n\nInterest runs daily."
      return "## Terms\n\n### Terms\n\nThe terms apply."
    })
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe("# Payment Demand\n\n## Demand for payment\n\nPay the sum due.\n\n### Interest\n\nInterest runs daily.\n\n## Terms\n\nThe terms apply.\n")
    expect(styledParagraphs(directory)).toEqual([
      ["Payment Demand", "Heading1", "320"],
      ["Demand for payment", "Heading2", "280"], ["Pay the sum due.", undefined, undefined],
      ["Interest", "Heading3", "220"], ["Interest runs daily.", undefined, undefined],
      ["Terms", "Heading2", "280"], ["The terms apply.", undefined, undefined],
    ])
  }, 60_000)

  test("a plan saved before labels were planned unheaded keeps them as visible headings and tells the drafter so", async () => {
    const { directory, state } = matter()
    writeLeanPlan(directory, [{ title: "Notice letter", purpose: "", parts: [{ heading: "Background", scope: "" }, { heading: "Notice", scope: "" }] }], outlineBasis(directory), NOT_STATED)
    const model = scripted(text => text.includes("part 1 of 2") ? "## Background\n\nThe parties signed the terms." : "## Notice\n\nNotice is given in writing.")
    const result = await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})
    expect(result.status).toBe("done")
    expect(model.prompts[0]!.text).toContain("The renderer prints that first line as the part's heading. Do not repeat it in the body.")
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe("# Notice letter\n\n## Background\n\nThe parties signed the terms.\n\n## Notice\n\nNotice is given in writing.\n")
    expect(readParagraphs(directory, "d1").map(row => row.text)).toEqual(["Notice letter", "Background", "The parties signed the terms.", "Notice", "Notice is given in writing."])
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
    expect(parsePlan("Acting for: not stated\n# Title ##\nForm: other\n## One\n- covers the start\n## Two").value?.documents)
      .toEqual([{ title: "Title", purpose: "", parts: [{ heading: "One", scope: "covers the start", display_heading: "" }, { heading: "Two", scope: "", display_heading: "" }], form: "other" }])
    // Several documents: each "# " starts one, its plain lines before the first
    // part are its purpose, and the same part heading may recur across documents.
    expect(parsePlan("Acting for: not stated\n# First\nWhat the first is for.\nForm: agreement\n## Terms\n# Second\nForm: advice\n## Terms\nCovers terms.").value?.documents).toEqual([
      { title: "First", purpose: "What the first is for.", parts: [{ heading: "Terms", scope: "", display_heading: "" }], form: "agreement" },
      { title: "Second", purpose: "", parts: [{ heading: "Terms", scope: "Covers terms.", display_heading: "" }], form: "advice" },
    ])
  })

  test("the outline records whom the lawyer acts for from one line before the first title", () => {
    const stated = parsePlan("Acting for: the buyer\n# Title\nForm: other\n## One")
    expect(stated).toEqual({ issues: [], value: { documents: [expect.objectContaining({ title: "Title" })], acting_for: { party: "the buyer", stated: true } } })
    expect(parsePlan("- **Acting for:** Not stated.\n# Title\nForm: other\n## One")).toEqual({ issues: [], value: expect.objectContaining({ acting_for: NOT_STATED }) })
    // Without the line the outline is retried; its value stands as not stated if the last attempt still lacks it.
    for (const reply of ["# Title\nForm: other\n## One", "# Title\nForm: other\nActing for: the buyer\n## One"]) {
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
    expect(model.prompts[2]!.text).toContain("<acting_for>\nThe lawyer acts for: the sender.\n</acting_for>")
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

  test("the part that opens the document is retried for an opening line repeating the printed title, naming the line; a later part, or a mention, keeps it", () => {
    const deed = { title: "Rent Deposit Deed", form: "other" as const, parts: [{ heading: "One", scope: "", display_heading: "" }, { heading: "Two", scope: "", display_heading: "" }] }
    const issue = (line: string) => `The part opens with the document title "Rent Deposit Deed". The title is printed once at the top of the document: keep the "## " line first and remove the line "${line}" after it.`
    for (const [reply, line] of [["## One\n\n**RENT DEPOSIT DEED**\n\nThis deed is made on [date].", "**RENT DEPOSIT DEED**"], ["**Rent deposit deed.**\n\nText.", "**Rent deposit deed.**"], ["## One\n\n### _Rent  Deposit_ `Deed`\n\nText.", "### _Rent  Deposit_ `Deed`"]])
      expect(checkSection(reply!, deed, 0).issues).toEqual([issue(line!)])
    expect(checkSection("## One\n\n**RENT DEPOSIT DEED**", deed, 0).issues).toEqual([issue("**RENT DEPOSIT DEED**"), 'The part "One" has no text under its heading.'])
    expect(checkSection("## One\n\nThis Rent Deposit Deed is made on [date].", deed, 0)).toEqual({ issues: [], value: "## One\n\nThis Rent Deposit Deed is made on [date]." })
    expect(checkSection("## One\n\n**1. Parties**\n\nText.\n\n**RENT DEPOSIT DEED**\n\nSigned as a deed.", deed, 0).issues).toEqual([])
    // On the last attempt the repeated title is left out rather than failing the stage.
    expect(checkSection("## One\n\n**RENT DEPOSIT DEED**\n\nText.", deed, 0, true)).toEqual({ issues: [issue("**RENT DEPOSIT DEED**")], value: "## One\n\nText." })
    // After a cover sheet the deed itself may restate its title.
    expect(checkSection("## Two\n\nRENT DEPOSIT DEED\n\nTHIS DEED is dated [date].", deed, 1)).toEqual({ issues: [], value: "## Two\n\nRENT DEPOSIT DEED\n\nTHIS DEED is dated [date]." })
    // A letter's planned title is not printed, so its text may open with it.
    for (const form of ["letter", "advice letter"] as const)
      expect(checkSection("## One\n\n**Rent Deposit Deed**\n\nWe write about the deposit.", { ...deed, form }, 0)).toEqual({ issues: [], value: "## One\n\n**Rent Deposit Deed**\n\nWe write about the deposit." })
  })

  test("a part that opens by repeating the heading printed for it is retried, naming the line, and on the last attempt kept without it", () => {
    const outline = { title: "Rent Deposit Deed", parts: [
      { heading: "Parties", scope: "", display_heading: "" }, { heading: "Deposit and interest", scope: "", display_heading: "2. Deposit and interest" }] }
    const reply = "## Deposit and interest\n\n**2. Deposit and interest**\n\n2.1 The deposit is held in a separate account."
    const issue = 'The part opens with its heading "2. Deposit and interest", which is printed from its "## " line: keep the "## " line first and remove the line "**2. Deposit and interest**" after it.'
    expect(checkSection(reply, outline, 1)).toEqual({ issues: [issue] })
    expect(checkSection(reply, outline, 1, true)).toEqual({ issues: [issue], value: "## Deposit and interest\n\n2.1 The deposit is held in a separate account." })
    expect(checkSection("## Deposit and interest\n\n2.1 The deposit of the 2. Deposit and interest clause.", outline, 1).issues).toEqual([])
    // A label printed as the part's heading: a plan saved before labels were unheaded, or headings shown.
    const shown = { title: "Notice letter", parts: [{ heading: "Background", scope: "" }, { heading: "Notice", scope: "" }] }
    expect(checkSection("## Notice\n\n### Notice\n\nNotice is given.", shown, 1).issues)
      .toEqual(['The part opens with its heading "Notice", which is printed from its "## " line: keep the "## " line first and remove the line "### Notice" after it.'])
    // A part whose label is not printed opens with its own heading.
    expect(checkSection("## Parties\n\n### Parties\n\nThe landlord and the tenant.", outline, 0).issues).toEqual([])
  })

  test("a part whose label is not printed is told to write its document's headings, following its source's unless changed; a letter's is told to write none", () => {
    const outline = { title: "Services agreement", form: "agreement" as const, parts: [{ heading: "Scope", scope: "", display_heading: "" }, { heading: "Duration", scope: "", display_heading: "" }] }
    const prompt = sectionPrompt(outline, 1, { source: "template.docx", parts: ["4 Duration"], left: ["3 Service Credits"] })
    expect(prompt).toContain('Write the headings the document itself needs in the part\'s text, in its own style, following the heading text and hierarchy of the template "template.docx" except where the outline or the instructions change them.')
    expect(prompt).toContain("Keep defined terms, numbering and cross-references consistent with the outline, the instructions and the parts already drafted.")
    expect(prompt.match(/numbering/g)).toHaveLength(1)
    expect(sectionPrompt({ ...outline, mode: "revise" }, 0)).toContain("following the heading text and hierarchy of the original except where the outline or the instructions change them.")
    expect(sectionPrompt(outline, 0)).toContain("Write the headings the document itself needs in the part's text, in its own style.\n")
    const letter = sectionPrompt({ ...outline, form: "letter" }, 0)
    expect(letter).toContain("That first line is an internal part label; the renderer omits it. Begin the body with the document text itself, without a replacement heading.")
    expect(letter).not.toContain("Write the headings")
  })

  test("assembly prints a part label only where the plan keeps it; a part without a printed heading moves its own headings up to the second level", () => {
    const parts = ["## One\n\n### 1. One\n\nText.\n\n#### 1.1 Detail\n\nMore.", "## Two\n\n### Sub\n\nMore."]
    const outline = { title: "T", parts: [{ heading: "One", scope: "" }, { heading: "Two", scope: "", display_heading: "Second" }] }
    expect(assembleDraft(outline, parts)).toBe("# T\n\n## One\n\n### 1. One\n\nText.\n\n#### 1.1 Detail\n\nMore.\n\n## Second\n\n### Sub\n\nMore.\n")
    const unheaded = { ...outline, parts: [{ ...outline.parts[0]!, display_heading: "" }, outline.parts[1]!] }
    expect(assembleDraft(unheaded, parts)).toBe("# T\n\n## 1. One\n\nText.\n\n### 1.1 Detail\n\nMore.\n\n## Second\n\n### Sub\n\nMore.\n")
    expect(assembleDraft(unheaded, parts, false)).toBe(`# T\n\n${parts.join("\n\n")}\n`)
    // The part's shallowest heading becomes a second-level heading, keeping the depth below it; a part with none is unchanged.
    expect(assembleDraft(unheaded, ["## One\n\n#### 1. One\n\nText.\n\n###### 1.1 Detail\n\nMore.", "## Two\n\nPlain."]))
      .toBe("# T\n\n## 1. One\n\nText.\n\n#### 1.1 Detail\n\nMore.\n\n## Second\n\nPlain.\n")
    expect(assembleDraft(unheaded, ["## One\n\nPlain text.", "## Two\n\nMore."])).toBe("# T\n\nPlain text.\n\n## Second\n\nMore.\n")
  })
})

// Each planned document's form, from its "Form:" line: page furniture follows
// it, and a letter reads as one letter, its planning labels unheaded.
describe("a planned document's form", () => {
  const FORMS = '"Form: agreement", "Form: advice", "Form: letter", "Form: advice letter", "Form: other"'
  const outline = (lines: string) => `Acting for: not stated\n# Notice letter\n${lines}\n## Opening\nWhy we write.\n## Terms\nDisplay heading: Terms of notice\nThe notice.\n## Closing\nDisplay heading: none\nWhat follows.`
  const shown = (lines: string) => parsePlan(outline(lines)).value?.documents[0]?.parts.map(part => part.display_heading)

  test("the outline prompt asks for each document's form and says labels are not printed, unless a letter shows them", () => {
    const prompt = stagePrompt(matter().directory, "d1.draft", { outline: true })
    expect(prompt).toContain('- Under each title write its form: "Form: agreement" for')
    expect(prompt).toContain('"Form: advice letter" for a letter giving the client legal advice, "Form: letter" for other correspondence addressed to a recipient')
    expect(prompt).toContain('otherwise "Form: other".')
    expect(prompt).toContain("Each part's ## label identifies the drafting job and is not printed; the part's text carries the document's own headings in its style. " +
      "To print a heading in a label's place, add Display heading: <publication heading> immediately under that label.")
    expect(prompt).toContain('For a letter whose text uses headings, write "Headings: shown" under its title to print its labels as its headings.')
    expect(prompt).toContain("print a heading they change as its part's Display heading.")
  })

  test("is read from one line under each title, ignoring case and list, bold or code marks, and an advice letter's hyphen or underscore", () => {
    for (const [line, form] of [["Form: agreement", "agreement"], ["- **Form:** Advice", "advice"], ["Form: `letter`.", "letter"], ["FORM:   other", "other"],
      ["Form: advice letter", "advice letter"], ["- **Form:** Advice-Letter.", "advice letter"], ["Form: `advice_letter`", "advice letter"]] as const) {
      const parsed = parsePlan(outline(line))
      expect(parsed.issues).toEqual([])
      expect(parsed.value?.documents[0]?.form).toBe(form)
    }
  })

  test("a missing or unknown form is retry feedback naming the document and the forms, and the document stands without one", () => {
    // Under a part the line is that part's scope, not the document's form.
    const missing = parsePlan("Acting for: not stated\n# Notice letter\n## Opening\nForm: letter")
    expect(missing.issues).toEqual([`The document "Notice letter" has no "Form:" line. Under its title, before its first "## " part, write one of ${FORMS}.`])
    expect(missing.value?.documents[0]).toEqual({ title: "Notice letter", purpose: "", parts: [{ heading: "Opening", scope: "Form: letter", display_heading: "" }] })
    expect(parsePlan(outline("Form:")).issues).toEqual([expect.stringContaining('The document "Notice letter" has no "Form:" line.')])
    const unknown = parsePlan(`${outline("Form: memo")}\n# Supply terms\nForm: agreement\n## Delivery`)
    expect(unknown.issues).toEqual([`The "Form:" line of "Notice letter" gives "memo", which is not one of the forms. Under its title, before its first "## " part, write one of ${FORMS}.`])
    expect(unknown.value?.documents.map(row => row.form)).toEqual([undefined, "agreement"])
    // Without a form the parts are unheaded too.
    expect(unknown.value?.documents[0]?.parts.map(part => part.display_heading)).toEqual(["", "Terms of notice", ""])
  })

  test("a missing Form line is retried in the outline session; an outline that never gives one is saved without a form, its labels unheaded", async () => {
    const part = (text: string) => text.includes("part 1 of 3") ? "The parties signed the terms." : text.includes("part 2 of 3") ? "Notice is given." : "We will wait."
    const formless = OUTLINE.replace("Form: letter\nHeadings: shown\n", "")
    for (const [given, form] of [[1, "letter"], [4, undefined]] as const) {
      const { directory, state } = matter()
      const model = scripted((text, index) => index < given ? formless : index === given && form ? OUTLINE.replace("Headings: shown\n", "") : part(text))
      expect((await runDraftStage(directory, state, { sessions: model.sessions }, () => {}, "d1", () => {})).status).toBe("done")
      expect(model.prompts[1]!.text).toContain(`The document "Notice letter" has no "Form:" line. Under its title, before its first "## " part, write one of ${FORMS}.`)
      expect(readLeanPlan(directory)?.documents[0]?.form).toBe(form)
      expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8").includes("Background")).toBe(false)
    }
  }, 60_000)

  test("every document's parts, of any form, are unheaded unless the outline shows them; its own Display heading lines stand", () => {
    for (const form of ["letter", "advice letter", "Advice-Letter", "agreement", "advice", "other"]) expect(shown(`Form: ${form}`)).toEqual(["", "Terms of notice", ""])
    for (const lines of ["Form: letter\nHeadings: shown", "- **Headings:** Shown.\nForm: Letter", "Form: advice letter\nHeadings: shown", "Form: agreement\nHeadings: shown"])
      expect(shown(lines)).toEqual([undefined, "Terms of notice", ""])
  })

  test("an unheaded letter, advice letter or not, goes through the draft stage's presentation path and keeps no planning label or title as a heading", async () => {
    const labels = ["Opening & purpose", "The client's position", "\u201cNext\u201d steps \u2014 *timing*"]
    for (const form of ["letter", "advice letter"] as const) {
      const { directory, state } = matter()
      const model = scripted((text, index) => {
        if (index === 0) return ["Acting for: not stated", "# Notice letter", `Form: ${form}`, ...labels.flatMap(label => [`## ${label}`, "What it covers."])].join("\n")
        if (text.includes("part 1 of 3")) return "We write about the terms."
        // A part may begin with its label line, which the check requires as written.
        if (text.includes("part 2 of 3")) return `## ${labels[1]}\n\n"Notice must be given in writing" [source: terms.docx].`
        return "We await your reply."
      })
      const result = await withFirm({}, () => runDraftStage(directory, dated(directory, state), { sessions: model.sessions }, () => {}, "d1", () => {}))
      expect(result.status).toBe("done")
      expect(readLeanPlan(directory)!.documents[0]).toMatchObject({ form, parts: labels.map(heading => ({ heading, display_heading: "" })) })
      expect(model.prompts.slice(1).every(prompt => prompt.text.includes("That first line is an internal part label; the renderer omits it."))).toBe(true)
      expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe([
        "# Notice letter", "", "We write about the terms.", "", "\"Notice must be given in writing\" [source: terms.docx].", "", "We await your reply.", "",
      ].join("\n"))
      const parts = unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))
      const body = strFromU8(parts["word/document.xml"]!)
      // Neither the planning labels nor the planned title print: the record keeps the title.
      expect(body).not.toContain('w:val="Heading2"')
      expect(body).not.toContain('w:val="Heading1"')
      const rows = readParagraphs(directory, "d1")
      expect(rows.map(row => [row.text, row.draft_part, row.draft_heading])).toEqual([
        ["We write about the terms.", 0, false], ["\"Notice must be given in writing\" [source: terms.docx].", 1, false], ["We await your reply.", 2, false],
      ])
      // A letter is numbered from its second page and carries the draft line on every page.
      expect(body).toContain("<w:titlePg")
      const header = ["Draft: 20 October 2026"]
      expect(furnitureLines(parts)).toEqual({ "header:default": header, "header:first": header, "footer:default": ["2"], "footer:first": [""] })
    }
  }, 60_000)

  test("an advice letter carries the firm's advice legend and the draft line from its first page, which has no page number", async () => {
    const { directory, state } = matter()
    const outline = "Acting for: not stated\n# Advice on the notice\nForm: advice letter\n## Advice\nThe advice."
    const model = scripted((_, index) => index === 0 ? outline : "We advise that notice be given in writing.")
    const firm = { firm: "Doe Partners LLP", adviceLegend: "Privileged and confidential", agreementLegend: "Subject to contract" }
    expect((await withFirm(firm, () => runDraftStage(directory, dated(directory, state), { sessions: model.sessions }, () => {}, "d1", () => {}))).status).toBe("done")
    const parts = unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))
    const header = ["Privileged and confidential", "Doe Partners LLP draft: 20 October 2026"]
    expect(furnitureLines(parts)).toEqual({ "header:default": header, "header:first": header, "footer:default": ["2"], "footer:first": [""] })
    expect(readParagraphs(directory, "d1").map(row => row.text)).toEqual(["We advise that notice be given in writing."])
  }, 60_000)

  test("what the plan asks of a document's furniture adds to the firm's settings in its body", async () => {
    const firm = { adviceLegend: "Privileged and confidential" }
    const header = ["Privileged and confidential", "Without prejudice", "Draft: 20 October 2026"]
    const cases: [PlannedDocument["furniture"], Record<string, string[]>][] = [
      [{ legend: "Without prejudice" }, { "header:default": header, "header:first": header, "footer:default": ["2"], "footer:first": [""] }],
      [{ legend: "", page_numbers: false }, { "header:default": ["Privileged and confidential", "Draft: 20 October 2026"] }],
    ]
    for (const [furniture, lines] of cases) {
      const { directory, state } = matter()
      dated(directory, state)
      writeLeanPlan(directory, [{ title: "Advice on the notice", purpose: "", parts: [{ heading: "Advice", scope: "" }], form: "advice letter", furniture }])
      expect((await withFirm(firm, () => acceptStageReply(directory, "d1.draft", "# Advice on the notice\n\n## Advice\n\nWe advise that notice be given in writing.\n"))).issues).toEqual([])
      expect(furnitureLines(unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx")))))).toEqual(lines)
    }
  }, 60_000)

  test("an agreement's body carries the firm's legend and a draft line dated by the run's start, outside its paragraphs", async () => {
    const { directory, state } = matter()
    const workspace = mkdtempSync(path.join(tmpdir(), "lean-furniture-prefs-"))
    directories.push(workspace)
    mkdirSync(path.join(workspace, ".preferences"))
    const root = process.env.WORKSPACE_ROOT
    process.env.WORKSPACE_ROOT = workspace
    try {
      writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify({ firm: "Doe & Partners LLP", agreementLegend: "Subject to contract", dateFormat: "day-month-year" }))
      writeLeanPlan(directory, [{ title: "Supply terms", purpose: "", parts: [{ heading: "Delivery", scope: "" }], form: "agreement" }])
      // A step restarted later in the run does not move the date.
      updateLeanStage(directory, writeLeanState(directory, { ...state, created_at: "2026-10-20T12:00:00.000Z" }), "d1.draft", { status: "running", started_at: "2026-10-22T12:00:00.000Z" })
      expect((await acceptStageReply(directory, "d1.draft", "# Supply terms\n\n## Delivery\n\nGoods are delivered on the agreed date.\n")).issues).toEqual([])
      const parts = unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))
      const header = strFromU8(parts["word/header1.xml"]!)
      expect(header).toContain("Subject to contract")
      expect(header).toContain("Doe &amp; Partners LLP draft: 20 October 2026")
      expect(strFromU8(parts["word/footer1.xml"]!)).toContain(" PAGE ")
      expect(readParagraphs(directory, "d1").map(row => row.text)).toEqual(["Supply terms", "Delivery", "Goods are delivered on the agreed date."])
    } finally {
      if (root === undefined) delete process.env.WORKSPACE_ROOT
      else process.env.WORKSPACE_ROOT = root
    }
  }, 60_000)

  test("a document drafted from a Word template takes its clause numbering and page set-up", async () => {
    const { directory } = matter()
    // An A4 precedent whose clauses are numbered by its own Word list.
    const precedent = unzipSync((await renderMarkdownToDocx("# Precedent terms\n\n## 1. Term\n\n1.1 The term is one year.")).bytes)
    const a4 = strFromU8(precedent["word/document.xml"]!).replace(/<w:pgSz [^>]*\/>/, '<w:pgSz w:w="11906" w:h="16838"/>')
    writeFileSync(path.join(directory, "precedent.docx"), zipSync({ ...precedent, "word/document.xml": strToU8(a4) }))
    writeLeanPlan(directory, [{ title: "Supply terms", purpose: "", parts: [{ heading: "Terms", scope: "" }], form: "agreement", template: "precedent.docx" }])
    expect((await acceptStageReply(directory, "d1.draft", "# Supply terms\n\n## Terms\n\n### 1. Definitions\n\n1.1 Terms have their defined meanings.\n\n### 2. Supply\n\n2.1 The Supplier shall deliver the Goods.\n")).issues).toEqual([])
    const parts = unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx"))))
    expect(strFromU8(parts["word/document.xml"]!)).toContain('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(strFromU8(parts["word/numbering.xml"]!)).toBe(strFromU8(precedent["word/numbering.xml"]!))
    expect(readParagraphs(directory, "d1").map(shownText)).toEqual(["Supply terms", "Terms", "1. Definitions", "1.1 Terms have their defined meanings.", "2. Supply", "2.1 The Supplier shall deliver the Goods."])
  }, 60_000)

  test("a saved plan reads with or without a form; an edited form must be one of the forms, and it keys the draft", () => {
    const { directory } = matter()
    writeLeanPlan(directory, [{ title: "Notice letter", purpose: "", parts: [{ heading: "Notice", scope: "" }] }])
    const plan = readLeanPlan(directory)!
    expect(plan.documents[0]).not.toHaveProperty("form")
    const legacy = stageInputSha(directory, "d1.draft")
    const edited = (form: unknown) => writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, documents: [{ ...plan.documents[0]!, form }] }))
    edited("letter")
    expect(readLeanPlan(directory)?.documents[0]?.form).toBe("letter")
    expect(stageInputSha(directory, "d1.draft")).not.toBe(legacy)
    for (const form of ["Letter", "memo", ""]) {
      edited(form)
      expect(readLeanPlan(directory)).toBeUndefined()
      expect(leanPlanIssue(directory)).toBe("document 1 is not a valid planned document")
    }
  })
})

// What a request asks of one document's page furniture, from optional lines
// under its title: they add to the firm's settings for that document only.
describe("a document's furniture lines", () => {
  const REQUEST = "Draft a short notice letter marked Without prejudice save as to costs, and leave out the page numbers."
  const outline = (...lines: string[]) => ["Acting for: not stated", "# Notice letter", "Form: letter", ...lines, "## Opening", "Why we write."].join("\n")
  const UNGROUNDED = 'The "Legend:" line of "Notice letter" gives "Strictly private", but the request does not use those words. Write a "Legend:" line only when the request ' +
    'asks for a document\'s header legend, in the request\'s own words; otherwise leave the line out.'

  test("the outline prompt asks for them only when the request asks, the legend in its own words", () => {
    expect(stagePrompt(matter().directory, "d1.draft", { outline: true })).toContain('- Only when the request asks for a page header legend on a document, write under its title ' +
      '"Legend: <the legend in the request\'s own words>"; only when it asks for a document without page numbers, "Page numbers: none".')
  })

  test("are read under a title, ignoring case, list, emphasis or code marks, wrapping quotes and a trailing full stop, a legend in the request's words, the first of each standing", () => {
    for (const [lines, furniture] of [
      [["Legend: Without prejudice save as to costs"], { legend: "Without prejudice save as to costs" }],
      [['- **Legend:** "without prejudice   save as to costs."'], { legend: "without prejudice save as to costs" }],
      [["*Legend:* WITHOUT PREJUDICE SAVE AS TO COSTS", "Legend: Strictly private"], { legend: "WITHOUT PREJUDICE SAVE AS TO COSTS" }],
      [["Legend: Without prejudice save as to costs", "Page numbers: none"], { legend: "Without prejudice save as to costs", page_numbers: false }],
      [["- **Page numbers:** `None`."], { page_numbers: false }],
    ] as const) {
      const parsed = parsePlan(outline(...lines), REQUEST)
      expect(parsed.issues).toEqual([])
      expect(parsed.value?.documents[0]?.furniture).toEqual(furniture)
    }
    // Without the lines, or with them under a part, the document asks nothing of its furniture.
    expect(parsePlan(outline(), REQUEST).value?.documents[0]).not.toHaveProperty("furniture")
    const under = parsePlan(`${outline()}\nLegend: none`, REQUEST)
    expect(under.value?.documents[0]).not.toHaveProperty("furniture")
    expect(under.value?.documents[0]?.parts[0]?.scope).toBe("Why we write. Legend: none")
  })

  test("a legend the request does not use as written is retry feedback and is dropped, the other lines kept", () => {
    const parsed = parsePlan(outline("Legend: Strictly private", "Page numbers: none"), REQUEST)
    expect(parsed.issues).toEqual([UNGROUNDED])
    expect(parsed.value?.documents[0]?.furniture).toEqual({ page_numbers: false })
    // Its words as the request writes them, not a party name's folded forms or part of a word.
    for (const legend of ["The without prejudice save as to costs", "Without prejudice, save as to costs", "Without prejudice save as to cost"])
      expect(parsePlan(outline(`Legend: ${legend}`), REQUEST).issues).toEqual([UNGROUNDED.replace("Strictly private", legend)])
    expect(parsePlan(outline("Legend: Without prejudice – save as to costs"), "Mark it Without prejudice - save as to costs.").issues).toEqual([])
  })

  test("a request never removes the firm's legend or draft line: those lines ask nothing", () => {
    const parsed = parsePlan(outline("Legend: none", "Draft line: none"), "Draft a short notice letter with no legend and no draft line.")
    expect(parsed.issues).toEqual([])
    expect(parsed.value?.documents[0]).not.toHaveProperty("furniture")
  })

  test("a page numbers value other than none is retry feedback naming the only value, and is dropped", () => {
    const parsed = parsePlan(outline("Page numbers: bottom right"), REQUEST)
    expect(parsed.issues).toEqual([
      'The "Page numbers:" line of "Notice letter" gives "bottom right", but the only value it takes is "Page numbers: none". Write it only when the request asks for the document without page numbers; otherwise leave the line out.',
    ])
    expect(parsed.value?.documents[0]).not.toHaveProperty("furniture")
  })

  test("an ungrounded legend is retried in the outline session; one that stands to the last attempt is dropped and the firm's legend kept", async () => {
    const part = "We advise that notice be given in writing."
    const reply = (legend: string) => `Acting for: not stated\n# Advice on the notice\nForm: advice letter\nLegend: ${legend}\n## Advice\nThe advice.`
    for (const [given, requested] of [[1, ["Without prejudice save as to costs"]], [4, []]] as const) {
      const { directory, state } = matter()
      writeLeanText(directory, "request.md", REQUEST)
      const model = scripted((text, index) => text.includes("Draft part") ? part : reply(index < given ? "Strictly private" : "Without prejudice save as to costs"))
      const result = await withFirm({ adviceLegend: "Privileged and confidential" }, () => runDraftStage(directory, dated(directory, state), { sessions: model.sessions }, () => {}, "d1", () => {}))
      expect(result.status).toBe("done")
      expect(model.prompts.filter(prompt => prompt.text.includes(UNGROUNDED.replace("Notice letter", "Advice on the notice")))).toHaveLength(given === 1 ? 1 : LEAN_MAX_RETRIES)
      expect(readLeanPlan(directory)?.documents[0]?.furniture).toEqual(given === 1 ? { legend: requested[0] } : undefined)
      expect(furnitureLines(unzipSync(new Uint8Array(readFileSync(leanFile(directory, "d1/body.docx")))))["header:default"]).toEqual(["Privileged and confidential", ...requested, "Draft: 20 October 2026"])
    }
  }, 60_000)

  test("a saved plan's furniture is checked as a lawyer may edit it, keys the draft and is not drafted", () => {
    const { directory } = matter()
    writeLeanPlan(directory, [{ title: "Notice letter", purpose: "", parts: [{ heading: "Notice", scope: "" }], form: "letter" }])
    const plan = readLeanPlan(directory)!
    const unedited = stageInputSha(directory, "d1.draft")
    const edited = (furniture: unknown) => writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, documents: [{ ...plan.documents[0]!, furniture }] }))
    for (const furniture of [{ legend: "" }, { page_numbers: false }, { legend: "Without prejudice", page_numbers: false }] satisfies PlannedDocument["furniture"][]) {
      edited(furniture)
      expect(readLeanPlan(directory)?.documents[0]?.furniture).toEqual(furniture)
    }
    expect(stageInputSha(directory, "d1.draft")).not.toBe(unedited)
    // The drafter's outline leaves it out: the header is rendered, not drafted.
    expect(operativeDocument(readLeanPlan(directory)!.documents[0])).not.toHaveProperty("furniture")
    for (const furniture of [null, [], "none", { legend: 1 }, { page_numbers: "none" }]) {
      edited(furniture)
      expect(readLeanPlan(directory)).toBeUndefined()
      expect(leanPlanIssue(directory)).toBe("document 1 is not a valid planned document")
    }
  })
})
