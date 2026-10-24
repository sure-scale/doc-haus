import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { advisePrompt } from "./advise"
import { checkChanges } from "./change-checks"
import { retypedLabelIssues, withoutListLabel, type ParagraphRow } from "./checks"
import { REQUEST_SOURCE, openItem } from "./quotes"
import { changesRepairSections } from "./repairs"
import { replaceStageOpenItems, writeLeanText, type Question } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

// Two paragraphs of a natively numbered body, as the lean rows record them:
// the list number Word prints is the label, not the text.
function numbered(directory: string): ParagraphRow[] {
  const rows = [
    { label: "2.", text: "Payment" },
    { label: "2.1", text: "The Customer shall pay within thirty days." },
  ].map((row, index) => ({ pid: `p${String(index + 1).padStart(12, "0")}`, sha256: paragraphSha256(row.text), ...row }))
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(rows))
  return rows
}

// A replacement may restate the number Word prints before its paragraph; only
// that number is the list's to print. Another number typed in its place is the
// model renumbering by hand, which the list would print beside its own label.
test("amendment text drops only its paragraph's own list label, in the forms Word prints it", () => {
  expect(withoutListLabel("2.2 Payment is made by card.", "2.2")).toBe("Payment is made by card.")
  expect(withoutListLabel("2.2\tPayment is made by card.", "2.2")).toBe("Payment is made by card.")
  expect(withoutListLabel("Article 1 Definitions", "Article 1")).toBe("Definitions")
  expect(withoutListLabel("3.2 Payment is made by card.", "2.2")).toBe("3.2 Payment is made by card.")
  expect(withoutListLabel("Article 2 Definitions", "Article 1")).toBe("Article 2 Definitions")
  expect(withoutListLabel("2.2", "2.2")).toBe("2.2")
})

test("a different number typed where the list prints one is named for a retry", () => {
  expect(retypedLabelIssues("amendments[0].updated", "3.2 Payment is made by card.", "2.2"))
    .toEqual([expect.stringContaining('amendments[0].updated begins with "3.2"')])
  expect(retypedLabelIssues("amendments[0].updated", "Article 2 Definitions", "Article 1")).toHaveLength(1)
  expect(retypedLabelIssues("amendments[0].updated", "(c) the third limb.", "(b)")).toHaveLength(1)
  expect(retypedLabelIssues("amendments[0].updated", "2.2 Payment is made by card.", "2.2")).toEqual([])
  expect(retypedLabelIssues("amendments[0].updated", "30 days after delivery.", "2.2")).toEqual([])
  expect(retypedLabelIssues("amendments[0].updated", "3.2 Payment is made by card.", undefined)).toEqual([])
})

test("a finding quoting a clause with the number its reader was shown names that clause for repair", () => {
  const directory = matter()
  const rows = numbered(directory)
  const item = openItem("draft", "quote_miss", "The payment term is not in the source.", { document: "d1", quote: "2.1 The Customer shall pay within" })
  replaceStageOpenItems(directory, "draft", "d1", [item])
  expect(changesRepairSections(directory, "d1").join("\n")).toContain(`${item.id} [${rows[1]!.pid}]`)
})

test("an adviser is shown the affected paragraph with its list number, as the drafting stages show it", () => {
  const directory = matter()
  const rows = numbered(directory)
  const question: Question = { id: "q1", stage: "changes", question: "How long may the Customer take to pay?", context: "", pid: rows[1]!.pid, document: "d1", status: "open" }
  expect(advisePrompt(directory, question)).toContain("<affected_paragraph>\n2.1 The Customer shall pay within thirty days.")
})

// The change stage retries a replacement that types a different number in
// place of the one Word prints, rather than shipping both numbers side by side.
test("a change that retypes another number before a Word-numbered paragraph is retried", () => {
  const directory = matter()
  const rows = numbered(directory)
  const amend = (updated: string | string[], operation?: string) =>
    JSON.stringify({ amendments: [{ pid: rows[1]!.pid, ...(operation ? { operation } : {}), updated, reason: "Review finding.", sources: [{ source_id: REQUEST_SOURCE, quote: "Draft a short supply note" }] }], questions: [] })
  expect(checkChanges(directory, amend("3.1 The Customer shall pay within ten days."), rows, { document: "d1" }).issues.join(" ")).toContain('begins with "3.1"')
  expect(checkChanges(directory, amend(["3.1 The Customer shall pay within ten days.", "Late sums bear interest."], "split"), rows, { document: "d1" }).issues.join(" ")).toContain('begins with "3.1"')
  const own = checkChanges(directory, amend("2.1 The Customer shall pay within ten days."), rows, { document: "d1" })
  expect(own.issues).toEqual([])
  expect(own.changes?.amendments[0]?.updated).toBe("The Customer shall pay within ten days.")
})
