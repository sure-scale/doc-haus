import { afterEach, expect, test } from "bun:test"
import { writeLeanPlan } from "./plan"
import { acceptReviewReply, findingsText, type RaisedFinding } from "./review-findings"
import { cleanupMatters, matter } from "./test-support"

// One review finding names every paragraph it affects and says, apart from what
// is wrong, what those paragraphs should say instead.

afterEach(cleanupMatters)

const FIRST = "p111111111111", SECOND = "p222222222222"
const paragraphs = [FIRST, SECOND].map(pid => ({ pid, sha256: pid, text: "The supplier shall deliver the goods." }))

function planned() {
  const directory = matter()
  const part = [{ heading: "Delivery", scope: "Delivery" }]
  writeLeanPlan(directory, [
    { title: "Supply note", purpose: "Record the supply terms", file: "supply note.docx", parts: part },
    { title: "Covering letter", purpose: "Cover the note", file: "covering letter.docx", parts: part },
  ])
  return directory
}

test("a finding keeps every paragraph it names, a lone pid among them, and its instead", () => {
  const reply = JSON.stringify({ findings: [{ pids: [FIRST, `[${SECOND}]`], pid: FIRST, origin: "drafting", finding: "Delivery is unclear.", instead: "Deliver within thirty days.", sources: [] }] })
  const accepted = acceptReviewReply(planned(), "d1", "review", reply, paragraphs)
  expect(accepted.issues).toEqual([])
  expect(accepted.record!.findings[0]).toMatchObject({ document: "d1", pids: [FIRST, SECOND], finding: "Delivery is unclear.", instead: "Deliver within thirty days." })
})

test("each unknown paragraph is reported and a finding for another document keeps none", () => {
  const directory = planned()
  const unknown = JSON.stringify({ findings: [{ pids: [FIRST, "p999999999999"], origin: "drafting", finding: "Delivery is unclear.", sources: [] }] })
  expect(acceptReviewReply(directory, "d1", "review", unknown, paragraphs).issues).toEqual(["findings[0]: [p999999999999] is not a paragraph of the draft. Use only the [pid] values listed with the draft."])
  const other = JSON.stringify({ findings: [{ document: "d2", pids: [FIRST], origin: "request", finding: "The letter should state the delivery period.", sources: [] }] })
  expect(acceptReviewReply(directory, "d1", "review", other, paragraphs).record!.findings[0]!.pids).toBeUndefined()
})

test("another document's view of a finding shows its instead without this draft's paragraph pointers", () => {
  const row: RaisedFinding = { id: "f-0000000001", document: "d2", origin: "request", finding: `The letter contradicts [${FIRST}].`,
    instead: `State the period of [${FIRST}].`, sources: [], raised_in: "d1", stage: "review" }
  const [shown] = JSON.parse(findingsText([row], "d2"))
  expect(shown).toMatchObject({ finding: "The letter contradicts (earlier paragraph).", instead: "State the period of (earlier paragraph)." })
})
