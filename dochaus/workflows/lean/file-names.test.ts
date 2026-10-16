import { afterEach, describe, expect, test } from "bun:test"
import { readLeanDocuments } from "./deliverables"
import { parsePlan, readLeanPlan } from "./plan"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { readLeanText, writeLeanText } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, happyPath, matter, registering, scriptedModel, type Reply } from "./test-support"

// Output file names the request asks for: the outline records them on a
// "File:" line under each document's title, the plan carries them apart from
// the title, and delivery uses them. Real coordinator; only the model is scripted.

afterEach(cleanupMatters)

const REQUEST = "Draft a short supply note from the supply terms and save it as supply-note-final.docx."

/** The scripted drafter whose outline replies are rewritten by `outline`. */
const outlined = (outline: (reply: string) => string): Reply => {
  const draft = drafter(DRAFT)
  return (prompt, index) => {
    const reply = draft(prompt, index)
    return reply.startsWith("Acting for:") ? outline(reply) : reply
  }
}

describe("requested output file names", () => {
  test("a file name the request gives is planned apart from the title and names the delivered files", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", REQUEST)
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: outlined(reply => reply.replace("# Supply note", "# Supply note\nFile: `supply-note-final.docx`")) })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(readLeanPlan(directory)?.documents.map(row => [row.title, row.file])).toEqual([["Supply note", "supply-note-final.docx"]])
    expect(readLeanText(directory, "d1/draft.md")?.split("\n")[0]).toBe("# Supply note")
    expect(readLeanDocuments(directory)!.documents.map(row => row.deliverable)).toEqual([
      "supply-note-final.docx", "supply-note-final - redline.docx", "supply-note-final - drafting notes.docx", null,
    ])
  }, 120_000)

  test("a title that carries the file name is retried with the instruction to give it on a File line", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", REQUEST)
    const deliverables = registering()
    const embedded = outlined(reply => reply.replace("# Supply note", "# Supply note — `supply-note-final.docx`"))
    const fixed = outlined(reply => reply.replace("# Supply note", "# Supply note\nFile: supply-note-final.docx"))
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: (prompt, index) => (index === 0 ? embedded : fixed)(prompt, index) })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    const retry = model.prompts.filter(prompt => prompt.agent === AGENT.draft)[1]!.text
    expect(retry).toContain("carries a file name")
    expect(readLeanPlan(directory)?.documents.map(row => [row.title, row.file])).toEqual([["Supply note", "supply-note-final.docx"]])
    expect(deliverables.names).toContain("supply-note-final.docx")
    expect(deliverables.names.some(name => name.includes("`"))).toBe(false)
  }, 120_000)

  test("a file name the request does not give is retried and, if the outline keeps it, dropped", () => {
    const outline = "Acting for: not stated\n# Supply note\nFile: `notes-v2.docx`\n## Delivery"
    const parsed = parsePlan(outline, REQUEST)
    expect(parsed.issues.join(" ")).toContain('gives "notes-v2.docx", but the request gives no such file name')
    expect(parsed.value?.documents[0]).not.toHaveProperty("file")
    expect(parsePlan("Acting for: not stated\n# Supply note\nFile: drafts/supply-note-final.docx\n## Delivery", REQUEST).issues.join(" "))
      .toContain("not a plain file name")
    expect(parsePlan("Acting for: not stated\n# A\nFile: supply-note-final.docx\n## One\n# B\nFile: supply-note-final.docx\n## Two", REQUEST).issues.join(" "))
      .toContain("given to more than one document")
  })

  test("a title the outline keeps with Markdown code marks after the last retry is delivered without them", async () => {
    const directory = matter()
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: outlined(reply => reply.replace("# Supply note", "# `Supply` note")) })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes("contains Markdown code marks"))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanPlan(directory)?.documents[0]?.title).toBe("Supply note")
    expect(deliverables.names).toContain("Supply note.docx")
  }, 120_000)
})
