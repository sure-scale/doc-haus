import { Database } from "bun:sqlite"
import path from "node:path"
import { parsePlan, type PlannedPart } from "./plan"
import { templateContract } from "./template-contract"
import { draftBody } from "./template-coverage"
import { tokens } from "./template-structure"
import { matter } from "./test-support"

// Shared fixtures for the template coverage tests. The template is an invented
// services agreement: four numbered sections and two annexes with text of
// their own.

export const TEMPLATE = "services-template.docx"
export const TEMPLATE_TEXT = [
  "SERVICES AGREEMENT",
  "1. Scope of Work\nThe provider shall run the help desk described in Annex A with reasonable skill and care, using suitably trained staff, and shall report to the customer on open requests at the end of each month.",
  "2. Charges\nThe customer shall pay the charges set out in Annex B within thirty days after receiving a correct invoice. The charges exclude sales taxes, which the customer pays in addition at the rate then in force.",
  "3. Service Credits\nIf the provider misses a response time in any month, the customer may claim a credit against the next invoice equal to five per cent of the charges for that month. The provider shall apply the credit without deduction or delay. The credit is the customer's only remedy for the missed response time, and the provider shall report each miss in its monthly summary.",
  "4. Duration\nThis agreement starts on the date of signature and runs for two years, unless either party ends it earlier by giving the other ninety days' written notice.",
  "ANNEX A – HELP DESK SERVICES\nThe provider shall staff the help desk from nine to five on working days. Each request is logged on receipt, answered within one working day and resolved within five working days where it concerns a routine matter. The provider keeps the log for the customer to inspect and sends a summary of all requests handled at the end of each month.",
  "ANNEX B – CHARGING SCHEDULE\nThe monthly charge covers up to forty hours of help desk time. Further hours are charged at the hourly rate the parties agree in writing and are invoiced monthly in arrears. Travel and other expenses are charged only where the customer has approved them in writing before they are incurred, and are passed on at cost without any mark-up.",
].join("\n\n")

// A template with no headed parts: plain paragraphs of prose.
export const PLAIN = "plain-terms.docx"
export const PLAIN_TEXT = [
  "The provider shall run the help desk with reasonable skill and care and report to the customer each month.",
  "The customer shall pay the charges within thirty days after receiving a correct invoice.",
  "Either party may end the arrangement by giving the other ninety days' written notice.",
].join("\n\n")

export const PARTS = {
  scope: "## 1. Scope of Work\n\nThe provider shall run the help desk described in Annex A with reasonable skill and care and shall report to the customer on open requests each month.",
  charges: "## 2. Charges\n\nThe customer shall pay the charges set out in Annex B within thirty days after receiving a correct invoice, plus sales taxes.",
  credits: "## 3. Service Credits\n\nIf the provider misses a response time in any month, the customer may claim a credit of five per cent of the charges for that month against the next invoice. The provider applies the credit without delay and reports each miss. The credit is the customer's only remedy for the miss.",
  duration: "## 4. Duration\n\nThis agreement runs for two years from signature unless either party ends it earlier on ninety days' written notice.",
  annexA: "## Annex A – Help Desk Services\n\nThe provider shall staff the help desk from nine to five on working days. Each request is logged on receipt, answered within one working day and resolved within five working days where routine. The provider keeps the log for the customer to inspect and sends a monthly summary.",
  annexB: "## Annex B – Charging Schedule\n\nThe monthly charge covers up to forty hours of help desk time. Further hours are charged at the agreed hourly rate and invoiced monthly in arrears. Expenses are charged at cost, and only with the customer's prior written approval.",
}
export const TITLE = "# Help desk agreement"
export const AGREEMENT = [TITLE, ...Object.values(PARTS)].join("\n\n")

/** A matter holding the template and the plain-prose source beside the
 * default source, each indexed as one chunk per paragraph as ingest does. */
export function templateMatter() {
  const directory = matter()
  insertChunks(directory, TEMPLATE, paragraphChunks(TEMPLATE_TEXT))
  insertChunks(directory, PLAIN, paragraphChunks(PLAIN_TEXT))
  return directory
}

export function paragraphChunks(text: string) {
  return text.split("\n\n").reduce<{ rows: { section: string; text: string; start: number }[]; at: number }>((acc, block, index) => ({
    rows: [...acc.rows, { section: String(index + 1), text: block, start: acc.at }],
    at: acc.at + block.length + 2,
  }), { rows: [], at: 0 }).rows
}

export function insertChunks(directory: string, name: string, chunks: { section: string; text: string; start: number }[]) {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  const insert = db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)")
  chunks.forEach(chunk => insert.run(name, chunk.section, chunk.text, chunk.start, chunk.start + chunk.text.length))
  db.close()
}

export function heading(part: string) {
  return part.split("\n")[0]!
}

/** The outlined documents and the first one's contract, for a template indexed
 * in a fresh matter under `name` and an outline whose first title gets the
 * Template line. */
export function contractOf(text: string, outline: string, name = "template.docx") {
  const directory = matter()
  insertChunks(directory, name, paragraphChunks(text))
  const documents = parsePlan(outline.replace(/^(# .+)$/m, `$1\nTemplate: ${name}`)).value!.documents
  return { directory, documents, contract: templateContract(directory, documents, 0)! }
}

/** A draft body from Markdown parts, each placed at the outlined part whose
 * heading it reads as (numbering aside), or at none. */
export function bodyOf(parts: PlannedPart[], ...texts: string[]) {
  return draftBody(texts.map(text => {
    const at = parts.findIndex(part => tokens(part.heading).join(" ") === tokens(heading(text).replace(/^##\s+/, "")).join(" "))
    return at < 0 ? { text } : { part: at, text }
  }))
}
