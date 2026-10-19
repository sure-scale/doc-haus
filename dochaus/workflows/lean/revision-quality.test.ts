import { afterEach, expect, test } from "bun:test"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../lib/render"
import { readChangeList, readParagraphs, type ApplyResult } from "./checks"
import { checkChanges } from "./change-checks"
import { readLeanDocuments, recordLeanDocuments } from "./deliverables"
import { parsePlan, writeLeanPlan } from "./plan"
import { completedDocumentsText, stagePrompt } from "./prompts"
import { runLean, stageInputSha } from "./runner"
import { leanFile, readLeanJson, readLeanText, writeLeanText, writeLeanBytes } from "./state"
import { documentTemplate, templateContract } from "./template-contract"
import { writeLegalCompactionPrompt } from "../../plugin/legal/compaction"
import { AGENT, cleanupMatters, drafter, happyPath, matter, pidOf, scriptedModel, SOURCE, SUPPORT } from "./test-support"

afterEach(cleanupMatters)

const original = "# Supply note\n\n## Delivery\n\nThe supplier shall deliver the goods within twenty days of the order date.\n\n## Payment\n\nPayment is due on delivery.\n"
const revised = original.replace("twenty", "thirty")
const request = "Revise the supplied note. Set delivery to thirty days. Deliver the annotated comparison as supply-review.docx and a covering letter."
const reason = "The proposed longer delivery period gives the supplier additional fulfilment time; the payment term remains unchanged."

function outline(prompt: string, index: number) {
  const answer = drafter(revised)(prompt, index)
  return prompt.includes('Draft part ') ? answer : answer.replace("# Supply note", `# Supply note\nMode: revise\nDeliverable: redline\nCompare against: ${SOURCE}\nFile: supply-review.docx`)
}

function explanations(prompt: string) {
  return JSON.stringify({ amendments: [], explanations: [{ pid: pidOf(prompt, "The supplier shall deliver the goods within thirty days"), reason,
    sources: [{ source_id: "request", quote: "Set delivery to thirty days." }, { source_id: SOURCE, quote: "within twenty days of the order date" }] }], questions: [] })
}

test("explicit revision mode preserves its source foundation and requested artifact role", () => {
  const directory = matter(original)
  const parsed = parsePlan(`Acting for: not stated\n# Supply note\nMode: revise\nDeliverable: redline\nCompare against: ${SOURCE}\nFile: supply-review.docx\n## Delivery\nDelivery\n## Payment\nPayment`, request)
  expect(parsed.issues).toEqual([])
  expect(parsed.value!.documents[0]).toMatchObject({ mode: "revise", deliverable: "redline" })
  writeLeanPlan(directory, parsed.value!.documents)
  expect(documentTemplate(directory, parsed.value!.documents[0])?.source).toBe(SOURCE)
  expect(stagePrompt(directory, "d1.draft")).toContain("Preserve unaffected source wording")
})

test("initial revision explanations become native comments without fake amendments and the requested filename contains the redline", async () => {
  const directory = matter(original)
  writeLeanText(directory, "request.md", request)
  writeFileSync(path.join(directory, SOURCE), (await renderMarkdownToDocx(original)).bytes)
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: outline, [AGENT.review]: () => "No corrections needed.", [AGENT.changes]: explanations })
  const result = await runLean(directory, { sessions: model.port, registerDeliverable: async (dir, name, bytes) => { writeFileSync(path.join(dir, name), bytes); return name } })
  expect(result.status).toBe("completed")
  expect(readChangeList(directory, "d1")!.amendments).toEqual([])
  const apply = readLeanJson(directory, "d1/apply.json") as ApplyResult
  expect(apply.applied).toEqual([])
  expect(apply.comments!.attached).toHaveLength(1)
  expect(apply.comments!.attached[0]!.text).toContain(reason)
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(readFileSync(leanFile(directory, "d1/body-applied.docx")))
  const delivered = unzipSync(readFileSync(path.join(directory, "supply-review.docx")))
  expect(strFromU8(delivered["word/comments.xml"]!)).toContain("fulfilment time")
  expect(strFromU8(delivered["word/document.xml"]!)).toContain("<w:del")
  expect(strFromU8(delivered["word/document.xml"]!)).toContain("<w:ins")
  expect(readLeanDocuments(directory)!.documents.find(row => row.kind === "body")!.deliverable).toBe("supply-review - clean.docx")
}, 120_000)

test("explanation admission verifies evidence without changing the body", () => {
  const directory = matter()
  const paragraphs = [{ pid: "p111111111111", sha256: "initial", text: "The supplier shall deliver the goods." }]
  const content = (quote: string, pid = paragraphs[0]!.pid) => JSON.stringify({ amendments: [], explanations: [{ pid, reason: "Explains the initial change.", sources: [{ source_id: SOURCE, quote }] }], questions: [] })
  const valid = checkChanges(directory, content(SUPPORT[0]!.quote), paragraphs)
  expect(valid.issues).toEqual([])
  expect(valid.changes).toHaveProperty("explanations")
  expect(checkChanges(directory, content("This invented passage is absent."), paragraphs).issues.join(" ")).toContain("not in that document")
  expect(checkChanges(directory, content(SUPPORT[0]!.quote, "p222222222222"), paragraphs).issues.join(" ")).toContain("not a paragraph")
})

test("later documents retain the accepted rationale and rekey when only that rationale changes", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record terms", parts: [{ heading: "Delivery", scope: "Terms" }] }, { title: "Covering letter", purpose: "Explain revisions", parts: [{ heading: "Advice", scope: "Consequences" }] }])
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p111111111111", sha256: "final", text: "The revised delivery term." }]))
  const changes = (text: string) => writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [], explanations: [{ id: "i1", pid: "p111111111111", reason: text, sources: SUPPORT }], questions: [] }))
  changes(reason)
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [], unapplied: [], comments: { author: "Doc.Haus", attached: [], unresolved: [] } }))
  const text = completedDocumentsText(directory, "d2")
  expect(text).toContain(reason)
  expect(text).toContain(SUPPORT[0]!.quote)
  expect(text).not.toContain("p111111111111")
  const key = stageInputSha(directory, "d2.change-plan")
  changes("A corrected source-linked rationale.")
  expect(stageInputSha(directory, "d2.change-plan")).not.toBe(key)
  expect(stagePrompt(directory, "d2.draft")).toContain("A corrected source-linked rationale.")
})

test("dependent advice retains selected guidance conclusions and rekeys without reopening the earlier review", () => {
  const directory = matter()
  const guidance = { sources: SUPPORT, reason: "Supplied terms govern delivery advice." }
  const documents = [
    { title: "Supply note", purpose: "Record terms", parts: [{ heading: "Delivery", scope: "Terms" }], playbook_review: guidance },
    { title: "Customer advice", purpose: "Explain the proposed changes", parts: [{ heading: "Advice", scope: "Consequences" }] },
  ]
  writeLeanPlan(directory, documents)
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p111111111111", sha256: "final", text: "The revised delivery term." }]))
  const earlierKey = stageInputSha(directory, "d1.review-playbook")
  const adviceKey = stageInputSha(directory, "d2.change-plan")
  const conclusion = "The original omission leaves delivery timing unspecified; the proposed period supplies a clear deadline."
  writeLeanText(directory, "d1/review-playbook.md", `### Coverage\n[p111111111111] ${conclusion}\nSource: ${SOURCE}.`)
  expect(stageInputSha(directory, "d2.change-plan")).not.toBe(adviceKey)
  expect(stageInputSha(directory, "d1.review-playbook")).toBe(earlierKey)
  const context = completedDocumentsText(directory, "d2")
  expect(context).toContain(conclusion)
  expect(context).toContain("not approval")
  expect(context).not.toContain("p111111111111")
  expect(stagePrompt(directory, "d1.review-independent")).not.toContain(conclusion)
  const updatedKey = stageInputSha(directory, "d2.change-plan")
  writeLeanText(directory, "d1/review-playbook.md", `### Coverage\n[p111111111111] Updated source interpretation must be checked against final wording.`)
  expect(stageInputSha(directory, "d2.change-plan")).not.toBe(updatedKey)
  writeLeanPlan(directory, documents.map(row => ({ ...row, playbook_review: undefined })))
  expect(completedDocumentsText(directory, "d2")).not.toContain("Updated source interpretation")
})

for (const fixture of [
  { title: "Equipment lease", heading: "Rent", before: "Monthly rent is 2400 units.", after: "Monthly rent is 2700 units.", preserved: "Notices must be addressed to the operations manager.", instruction: "Change the monthly rent to 2700 units.", rationale: "Implements the instructed rent increase while preserving the notice recipient." },
  { title: "Software licence", heading: "Renewal", before: "The licence renews automatically for another year.", after: "Renewal requires both parties' written agreement.", preserved: "The licensee may retain its own records after expiry.", instruction: "Replace automatic renewal with renewal by written agreement of both parties.", rationale: "Removes automatic renewal as instructed; the records-retention provision is unaffected." },
]) test(`${fixture.title}: initial explanation and unaffected provisions survive the native apply and delivery path`, async () => {
  const source = `# ${fixture.title}\n\n## ${fixture.heading}\n\n${fixture.before}\n\n## Other terms\n\n${fixture.preserved}\n`
  const draft = source.replace(fixture.before, fixture.after)
  const directory = matter(source)
  const requested = `Revise the supplied original. ${fixture.instruction} Deliver its annotated comparison as revised-document.docx.`
  writeLeanText(directory, "request.md", requested)
  writeFileSync(path.join(directory, SOURCE), (await renderMarkdownToDocx(source)).bytes)
  const agent = drafter(draft)
  const model = scriptedModel({ ...happyPath, [AGENT.review]: () => "No further correction.", [AGENT.draft]: (prompt, index) => {
    const reply = agent(prompt, index)
    return prompt.includes("Draft part ") ? reply : reply.replace(`# ${fixture.title}`, `# ${fixture.title}\nMode: revise\nDeliverable: redline\nCompare against: ${SOURCE}\nFile: revised-document.docx`)
  }, [AGENT.changes]: prompt => JSON.stringify({ amendments: [], explanations: [{ pid: pidOf(prompt, fixture.after), reason: fixture.rationale, sources: [{ source_id: "request", quote: fixture.instruction }, { source_id: SOURCE, quote: fixture.before }] }], questions: [] }) })
  const result = await runLean(directory, { sessions: model.port, registerDeliverable: async (dir, name, bytes) => { writeFileSync(path.join(dir, name), bytes); return name } })
  expect(result.status).toBe("completed")
  expect(readParagraphs(directory, "d1", "paragraphs-applied.json").some(row => row.text === fixture.preserved)).toBe(true)
  const exported = unzipSync(readFileSync(path.join(directory, "revised-document.docx")))
  expect(strFromU8(exported["word/comments.xml"]!)).toContain(fixture.rationale)
  expect(strFromU8(exported["word/document.xml"]!)).toContain("<w:del")
}, 120_000)


test("generated companions cannot claim another document's explicitly requested filename", async () => {
  const directory = matter()
  writeLeanPlan(directory, [
    { title: "Equipment comparison", purpose: "Annotated revision", deliverable: "redline", file: "review.docx", parts: [{ heading: "Terms", scope: "Terms" }] },
    { title: "Client advice", purpose: "Requested advice", file: "Review - clean.docx", parts: [{ heading: "Advice", scope: "Advice" }] },
  ])
  for (const [name, value] of [["d1/body-applied.docx", 1], ["d1/redline.docx", 2], ["d2/body-applied.docx", 3], ["d2/redline.docx", 4]] as const)
    writeLeanBytes(directory, name, (await renderMarkdownToDocx(`Generated record ${value}.`)).bytes)
  const result = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: async (dir, name, bytes) => { writeFileSync(path.join(dir, name), bytes); return name } }, true)
  expect(result.documents.find(row => row.document === "d1" && row.kind === "redline")!.deliverable).toBe("review.docx")
  expect(result.documents.find(row => row.document === "d2" && row.kind === "body")!.deliverable).toBe("Review - clean.docx")
  expect(readFileSync(path.join(directory, "Review - clean.docx"))).toEqual(readFileSync(leanFile(directory, "d2/body-applied.docx")))
  expect(result.documents.find(row => row.document === "d1" && row.kind === "body")!.deliverable).toBe("review - clean (2).docx")
})


test("revised operative coverage cannot be satisfied by a sibling advice document", () => {
  const directory = matter("# Equipment lease\n\n## Rent\n\nMonthly rent is payable on the first day of each calendar month.\n\n## Notices\n\nWritten notices must be sent to the operations manager at the designated address.\n")
  const documents = [
    { title: "Revised lease", purpose: "Revise the original", mode: "new" as "new" | "revise", template: SOURCE, parts: [{ heading: "Rent", scope: "Rent and notices" }] },
    { title: "Advice", purpose: "Explain the original", template: SOURCE, parts: [{ heading: "Notices", scope: "Explain notices" }] },
  ]
  const before = templateContract(directory, documents, 0)!
  const notice = before.top.find(unit => unit.label.includes("Notices"))!
  expect(before.places.get(notice.index)?.kind).toBe("document")
  documents[0]!.mode = "revise"
  const after = templateContract(directory, documents, 0)!
  expect(after.places.get(notice.index)?.kind).toBe("parts")
})

test("accepted exact sections remain retrievable after a section-boundary compaction handoff", async () => {
  const directory = matter(original)
  writeLeanText(directory, "request.md", request)
  const draft = drafter(revised)
  let compacted = false
  const model = scriptedModel({ ...happyPath, [AGENT.review]: () => "No corrections.", [AGENT.changes]: explanations,
    [AGENT.draft]: (prompt, index) => {
      if (prompt.includes('Draft part 2 of')) {
        // Discard dependence on the previous reply: retrieve exact text using
        // only the durable pointer available to a compacted child session.
        const handoff: { context: string[]; prompt?: string } = { context: [] }
        writeLegalCompactionPrompt(directory, handoff)
        expect(handoff.prompt).toContain("draft-in-progress.md")
        expect(prompt).toContain(".dochaus/lean/d1/draft-in-progress.md")
        expect(readLeanText(directory, "d1/draft-in-progress.md")).toContain("within thirty days of the order date")
        expect(readLeanText(directory, "d1/draft-in-progress.md")).not.toContain("Payment is due")
        compacted = true
      } else if (prompt.includes('Draft part 1 of')) {
        expect(readLeanText(directory, "d1/draft-in-progress.md")).toBe("")
      }
      return draft(prompt, index)
    },
  })
  writeLeanText(directory, "d1/draft-in-progress.md", "STALE DRAFT FROM AN EARLIER ATTEMPT")
  const result = await runLean(directory, { sessions: model.port })
  expect(result.status).toBe("completed")
  expect(compacted).toBe(true)
  expect(readLeanText(directory, "d1/draft-in-progress.md")).toContain("Payment is due on delivery")
  expect(readLeanText(directory, "d1/draft-in-progress.md")).not.toContain("STALE")
}, 120_000)

test("package repair corrects only native explanations and refreshes downstream advice through its change plan", async () => {
  const directory = matter(original)
  writeLeanText(directory, "request.md", request)
  writeFileSync(path.join(directory, SOURCE), (await renderMarkdownToDocx(original)).bytes)
  const advice = "# Covering letter\n\n## Advice\n\nThe proposed revision extends the delivery period.\n"
  const draft = drafter(revised, advice)
  let bodyChanges = 0
  const model = scriptedModel({ ...happyPath,
    [AGENT.draft]: (prompt, index) => {
      const reply = draft(prompt, index)
      return prompt.includes("Draft part ") ? reply : reply.replace("# Supply note", `# Supply note\nMode: revise\nDeliverable: redline\nCompare against: ${SOURCE}\nFile: supply-review.docx`)
    },
    [AGENT.review]: () => "No operative correction.",
    [AGENT.changes]: prompt => {
      if (!/\[p[0-9a-f]{12}\] The supplier shall/.test(prompt)) return JSON.stringify({ amendments: [], explanations: [], questions: [] })
      bodyChanges += 1
      return JSON.stringify({ amendments: [], explanations: [{ pid: pidOf(prompt, "The supplier shall deliver the goods within thirty days"), reason: bodyChanges === 1 ? "Incorrectly describes the original delivery period." : reason,
        sources: [{ source_id: SOURCE, quote: "within twenty days of the order date" }, { source_id: "request", quote: "Set delivery to thirty days." }] }], questions: [] })
    },
    [AGENT["package-review"]]: (prompt, index) => {
      expect(prompt).toContain("explanation-only correction")
      return JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] } : { status: "changes-required", delivery: [], findings: [{ target: "body", document: "d1", pid: pidOf(prompt, "The supplier shall deliver the goods within thirty days"), detail: "Correct the initial explanation to state that the original period was twenty days and the proposed period is thirty days. The operative body is correct; do not change it." }], limitations: [] })
    },
  })
  const result = await runLean(directory, { sessions: model.port, registerDeliverable: async (dir, name, bytes) => { writeFileSync(path.join(dir, name), bytes); return name } })
  expect(result.status).toBe("completed")
  expect(bodyChanges).toBe(2)
  expect(readChangeList(directory, "d1")!.amendments).toEqual([])
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(readFileSync(leanFile(directory, "d1/body-applied.docx")))
  const exported = unzipSync(readFileSync(path.join(directory, "supply-review.docx")))
  expect(strFromU8(exported["word/comments.xml"]!)).toContain("additional fulfilment time")
  expect(strFromU8(exported["word/comments.xml"]!)).not.toContain("Incorrectly describes")
  // The covering letter is drafted once; the corrected explanation reaches it through its change plan.
  const later = model.prompts.filter(row => row.agent === AGENT.draft && row.text.includes('Draft part 1 of 1 of "Covering letter"'))
  expect(later).toHaveLength(1)
  expect(later[0]!.text).toContain("Incorrectly describes")
  const letterPlans = new Set(model.sessions.flatMap((session, index) => session.title === "Lean drafting: Covering letter: change-plan" ? [`ses_child${index + 1}`] : []))
  expect(model.prompts.filter(row => letterPlans.has(row.sessionID)).at(-1)!.text).toContain(reason)
}, 120_000)
