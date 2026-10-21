import { afterEach, describe, expect, test } from "bun:test"
import { sectionPrompt } from "./draft-sections"
import type { Amendment } from "../../lib/drafting-amendment"
import type { ParagraphRow } from "./checks"
import { parsePlan, writeLeanPlan } from "./plan"
import { documentTemplate, expectedWords, outlineCoverageIssues, partCarries, templateContract } from "./template-contract"
import { contractGaps, draftTemplateItems, rowsBody, templateItems } from "./template-coverage"
import { parseTemplateUnits, sourceText } from "./template-structure"
import { PARTS, PLAIN_TEXT, TEMPLATE, TEMPLATE_TEXT, TITLE, bodyOf, contractOf, heading, insertChunks, templateMatter, textParagraphs } from "./template-test-support"
import { cleanupMatters, matter } from "./test-support"

// Template structure, coverage contract and measurement on invented templates.

afterEach(cleanupMatters)

const OUTLINE = [TITLE, ...Object.values(PARTS).map(heading)].join("\n")
const gapsOf = (built: ReturnType<typeof contractOf>, ...texts: string[]) =>
  contractGaps(built.contract, bodyOf(built.documents[0]!.parts, ...texts)).map(gap => [gap.unit.label, gap.gap])

// Two dividers above numbered clauses and a recital before them.
const DIVIDED = [
  "MASTER SUPPLY TERMS",
  "BACKGROUND\nThe parties wish to set out the terms on which the supplier sells goods to the buyer from time to time.",
  "PART 1 – GENERAL",
  "1. Orders\nThe buyer places each order in writing, and the supplier confirms or declines each order within two working days of receiving it.",
  "2. Delivery\nThe supplier delivers the goods to the address in the order on the agreed date, and the risk in the goods passes on delivery.",
  "PART 2 – PAYMENT",
  "3. Price\nThe price of the goods is the price in the supplier's list on the date of the order, and it includes packaging and carriage to the address.",
  "4. Invoicing\nThe supplier invoices the buyer on delivery, and the buyer pays each correct invoice within thirty days of receiving it.",
].join("\n\n")

// Numbered definitions inside a clause.
const DEFINED = [
  "LICENCE TERMS",
  "1. Definitions\nIn these terms the following words have the meanings below.\n1.1 Affiliate: any entity that controls, is controlled by or is under common control with a party.\n1.2 \"Business Day\"\na day other than a weekend or a public holiday in the place where the licensee is established.",
  "2. Licence\nThe licensor grants the licensee a non-exclusive licence to use the licensed software for its internal business purposes during the term.",
  "3. Fees and Payment\nThe licensee pays the fees annually in advance within thirty days of each invoice, and late amounts carry interest at the agreed rate.",
].join("\n\n")

// Alternative wordings offered inside a clause.
const ALTERNATIVES = [
  "SERVICES TERMS",
  "1. Services\nThe supplier provides the services described in the order with reasonable care and skill at all times during the term.",
  "2. Liability",
  "Option A\nEach party's total liability under these terms is limited to the charges paid in the twelve months before the claim arose.",
  "Option B\nEach party's total liability under these terms is limited to twice the charges paid in the twelve months before the claim arose.",
  "3. Term\nThese terms run for one year and renew each year unless either party gives the other three months' written notice.",
].join("\n\n")

// Clause headings in sentence case on lines of their own, one of them a single
// word, then a schedule under a sentence-case label whose table holds a line
// that opens with a number.
const SENTENCE_CASE_CLAUSES = {
  scope: ["1  Scope of services", "1.1  The contractor shall clean the offices at the address in the schedule on each working day, using its own trained staff and its own equipment and materials."],
  waste: ["2  Waste",
    "2.1  The contractor shall remove all general waste from the offices each evening and shall separate paper, glass and plastics for recycling as the client's written waste policy requires.",
    "2.2  The contractor shall not leave waste in shared corridors or the loading bay overnight, and shall report a full or damaged bin to the client's facilities manager on the same day."],
  fees: ["3  Fees and invoicing",
    "3.1  The client shall pay the monthly fee within thirty days after receiving a correct invoice. The fee is fixed for the first year and covers all labour, equipment and cleaning materials.",
    "3.2  Additional cleaning the client asks for in writing is charged at the hourly rate in the schedule and invoiced with the next monthly fee, with a note of the hours worked."],
  ending: ["4  Ending the agreement",
    "4.1  Either party may end this agreement by giving the other two months' written notice. Either party may end it at once by notice if the other commits a serious breach and does not put it right within fourteen days.",
    "4.2  Ending this agreement does not affect rights that have already accrued, and the contractor shall return all keys and access cards within two working days."],
  liability: ["5  Liability and insurance",
    "5.1  The contractor is responsible for loss or damage its staff cause at the offices and shall keep public liability insurance of at least the amount in the schedule throughout the term, showing the client the certificate on request."],
  schedule: ["Schedule 1  Site and hours",
    "The offices are on the second floor of the client's building, and the contractor cleans them between six and nine in the evening on each working day.",
    "Term", "1 March 2027 to 28 February 2028", "Hourly rate", "Thirty pounds for each additional hour, invoiced monthly in arrears."],
}
const SENTENCE_CASE = ["Office cleaning agreement", ...Object.values(SENTENCE_CASE_CLAUSES).flat()].join("\n\n")

describe("template structure", () => {
  test("a template's numbered sections and annexes are its parts; prose with no headed lines has none", () => {
    const units = parseTemplateUnits(TEMPLATE_TEXT)
    expect(units.map(unit => unit.label)).toEqual([
      "SERVICES AGREEMENT", "1 Scope of Work", "2 Charges", "3 Service Credits", "4 Duration",
      "ANNEX A – HELP DESK SERVICES", "ANNEX B – CHARGING SCHEDULE",
    ])
    expect(units[0]!.title).toBe(true)
    expect(units.slice(1).every(unit => unit.parent === undefined)).toBe(true)
    expect(units.slice(5).every(unit => unit.words >= 40)).toBe(true)
    expect(parseTemplateUnits(PLAIN_TEXT)).toEqual([])
  })

  test("the source text is rebuilt from overlapping chunks without repeating the overlap", () => {
    const directory = matter()
    const overlap = 20
    const cut = TEMPLATE_TEXT.indexOf("2. Charges")
    insertChunks(directory, "overlap.docx", [
      { section: "1", text: TEMPLATE_TEXT.slice(0, cut), start: 0 },
      { section: "2", text: TEMPLATE_TEXT.slice(cut - overlap), start: cut - overlap },
    ])
    expect(sourceText(directory, "overlap.docx")).toBe(TEMPLATE_TEXT)
  })

  test("the source text keeps the next heading whole after a chunk stored with an extra newline past its section", () => {
    // Chunks ingested before the chunker stopped counting one newline past a
    // split section: the next chunk's text replaces that newline.
    const directory = matter()
    const cut = TEMPLATE_TEXT.indexOf("2. Charges")
    insertChunks(directory, "older.docx", [
      { section: "1", text: `${TEMPLATE_TEXT.slice(0, cut)}\n`, start: 0 },
      { section: "2", text: TEMPLATE_TEXT.slice(cut), start: cut },
    ])
    expect(sourceText(directory, "older.docx")).toBe(TEMPLATE_TEXT)
  })

  test("a divider above numbered clauses is not a part, a recital before them is, and each clause stays a part of its own", () => {
    const units = parseTemplateUnits(DIVIDED)
    expect(units.map(unit => unit.label)).toEqual(["MASTER SUPPLY TERMS", "BACKGROUND", "1 Orders", "2 Delivery", "3 Price", "4 Invoicing"])
    expect(units.slice(1).every(unit => unit.parent === undefined)).toBe(true)
    const outline = ["# Supply terms", "## Background", "## 1. Orders", "## 2. Delivery", "## 3. Price", "## 4. Invoicing"].join("\n")
    const built = contractOf(DIVIDED, outline)
    expect(built.contract.unaccounted).toEqual([])
    const parts = DIVIDED.split("\n\n").slice(1).filter(block => !block.startsWith("PART")).map(block => `## ${block}`)
    expect(gapsOf(built, ...parts)).toEqual([])
    expect(gapsOf(built, ...parts.filter(part => !part.startsWith("## 3. Price")))).toEqual([["3 Price", "missing"]])
  })

  test("numbered paragraphs inside a schedule after the clauses belong to the schedule", () => {
    const template = [
      "SUPPORT TERMS",
      "1. Services\nThe supplier provides the support services in the schedule with reasonable skill and care throughout the term.",
      "2. Payment\nThe customer pays each invoice within thirty days of receipt, and late amounts carry interest at the agreed rate.",
      "SCHEDULE 1 – SERVICE LEVELS",
      "1. Response Times\nThe supplier answers each call within one working hour and attends the site within one working day of a reported fault.",
      "2. Reporting\nThe supplier reports on every call and visit at the end of each month, with the time taken to resolve each fault.",
    ].join("\n\n")
    const units = parseTemplateUnits(template)
    expect(units.map(unit => [unit.label, unit.parent])).toEqual([
      ["SUPPORT TERMS", undefined], ["1 Services", undefined], ["2 Payment", undefined], ["SCHEDULE 1 – SERVICE LEVELS", undefined],
      ["1 Response Times", 3], ["2 Reporting", 3],
    ])
  })

  test("numbered definitions are not parts", () => {
    expect(parseTemplateUnits(DEFINED).map(unit => unit.label)).toEqual(["LICENCE TERMS", "1 Definitions", "2 Licence", "3 Fees and Payment"])
  })

  test("alternative wordings inside a clause belong to it, and one left out on purpose is not expected", () => {
    const units = parseTemplateUnits(ALTERNATIVES)
    expect(units.map(unit => [unit.label, unit.parent])).toEqual([
      ["SERVICES TERMS", undefined], ["1 Services", undefined], ["2 Liability", undefined], ["Option A", 2], ["Option B", 2], ["3 Term", undefined],
    ])
    const outline = ["# Services terms", 'Left out: "Option B" – the parties chose the first liability cap.', "## 1. Services", "## 2. Liability", "## 3. Term"].join("\n")
    const built = contractOf(ALTERNATIVES, outline)
    const liability = built.contract.units[2]!
    expect(expectedWords(built.contract, liability)).toBe(liability.words - built.contract.units[4]!.words)
    const carries = partCarries(built.contract, 3)[1]!
    expect(carries.carried.map(unit => unit.label)).toEqual(["2 Liability"])
    expect(sectionPrompt(built.documents[0]!, 1, { source: "template.docx", parts: carries.carried.map(unit => unit.label), left: carries.left.map(unit => unit.label) }))
      .toContain('Leave out "Option B" of the template "template.docx", as the outline records.')
  })

  test("a clause heading in sentence case on its own line opens its own part, so each clause is measured by its own text", () => {
    const clauses = Object.values(SENTENCE_CASE_CLAUSES)
    const wordsOf = (lines: string[]) => lines.join(" ").match(/[\p{L}\p{N}]+/gu)!.length
    const outline = ["# Office cleaning agreement", ...clauses.map(([title]) => `## ${title!.replace(/^(\d+) /, "$1.").replace("Schedule 1 ", "Schedule 1 –")}`)].join("\n")
    const built = contractOf(SENTENCE_CASE, outline)
    expect(built.contract.top.map(unit => [unit.label, unit.words])).toEqual(clauses.map(([title, ...body]) => [title!.replace(/^(\d+) {2}/, "$1 "), wordsOf(body)]))
    expect(built.contract.unaccounted).toEqual([])
    const drafted = clauses.map(([, ...body], at) => [`## ${built.documents[0]!.parts[at]!.heading}`, ...body].join("\n\n"))
    expect(gapsOf(built, ...drafted)).toEqual([])
    const cut = drafted.map((text, at) => at === 1 ? `## ${built.documents[0]!.parts[1]!.heading}\n\n2.1 The contractor clears the bins each evening.` : text)
    expect(gapsOf(built, ...cut)).toEqual([["2 Waste", "thin"]])
    // The schedule's label names it under a reworded title.
    const rename = (text: string) => text.replace("Schedule 1 – Site and hours", "Schedule 1 – Premises and cleaning times")
    const renamed = contractOf(SENTENCE_CASE, rename(outline))
    expect(renamed.contract.unaccounted).toEqual([])
    expect(gapsOf(renamed, ...drafted.map(rename))).toEqual([])
  })
})

describe("coverage contract", () => {
  test("only the explicit Template line makes a document templated", () => {
    const directory = templateMatter()
    const documents = parsePlan(OUTLINE).value!.documents
    expect(documents[0]!.template).toBeUndefined()
    expect(documentTemplate(directory, documents[0])).toBeUndefined()
    expect(templateContract(directory, documents, 0)).toBeUndefined()
    expect(outlineCoverageIssues(directory, documents)).toEqual([])
  })

  test("a part is carried by the part whose heading or scope names it, or left out with a reason", () => {
    const outline = [TITLE, 'Left out: "3. Service Credits" – no response times are agreed.', heading(PARTS.scope),
      "## Fees", "The template's charges clause and its charging schedule.", heading(PARTS.duration), heading(PARTS.annexA)].join("\n")
    const built = contractOf(TEMPLATE_TEXT, outline)
    const places = Object.fromEntries([...built.contract.places].map(([index, place]) => [built.contract.units[index]!.label, place]))
    expect(places).toEqual({
      "1 Scope of Work": { kind: "parts", parts: [0] },
      "2 Charges": { kind: "parts", parts: [1] },
      "3 Service Credits": { kind: "left_out", line: "3. Service Credits", reason: "no response times are agreed." },
      "4 Duration": { kind: "parts", parts: [2] },
      "ANNEX A – HELP DESK SERVICES": { kind: "parts", parts: [3] },
      "ANNEX B – CHARGING SCHEDULE": { kind: "parts", parts: [1] },
    })
    expect(outlineCoverageIssues(built.directory, built.documents)).toEqual([])
    expect(partCarries(built.contract, 4).map(row => row.carried.map(unit => unit.label)))
      .toEqual([["1 Scope of Work"], ["2 Charges", "ANNEX B – CHARGING SCHEDULE"], ["4 Duration"], ["ANNEX A – HELP DESK SERVICES"]])
  })

  test("admission names unaccounted parts, Left out lines that name nothing, and Left out lines with no reason", () => {
    const outline = [TITLE, 'Left out: "Warranties" – not offered.', 'Left out: "3. Service Credits"', heading(PARTS.scope), "## Fees",
      heading(PARTS.duration), heading(PARTS.annexA), heading(PARTS.annexB)].join("\n")
    const built = contractOf(TEMPLATE_TEXT, outline)
    expect(outlineCoverageIssues(built.directory, built.documents)).toEqual([
      'The "Left out:" line "Warranties" of "Help desk agreement" names no part of its template "template.docx". Name the template part by its heading as the template writes it.',
      'The "Left out:" line "3. Service Credits" of "Help desk agreement" gives no reason. Say after the part\'s name why it is left out.',
      'The outline of "Help desk agreement" does not account for these parts of its template "template.docx": "2 Charges". ' +
      'Name each in the "## " heading or scope line of the part that carries it, or, for a part left out on purpose, add a line "Left out: "<template part>" – <reason>" under the title.',
    ])
  })

  test("the outline's Template and Left out lines are read with or without quotes", () => {
    const parsed = parsePlan(`${TITLE}\nTemplate: "${TEMPLATE}"\nLeft out: Service Credits – not agreed.\nLeft out: "Option B": the first cap applies.\n## 1. Scope of Work\n# Covering letter\n## Enclosure`)
    expect(parsed.value?.documents.map(document => document.template)).toEqual([TEMPLATE, undefined])
    expect(parsed.value?.documents[0]!.left_out).toEqual([{ part: "Service Credits", reason: "not agreed." }, { part: "Option B", reason: "the first cap applies." }])
    expect(parsed.value?.documents[1]!.left_out).toBeUndefined()
  })

  test("a part is not carried by another document drafted from the same template only because its title names it", () => {
    const outline = [TITLE, ...[PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration, PARTS.annexA].map(heading),
      "# Charging schedule", `Template: template.docx`, "## Monthly charge", "## Further hours and expenses"].join("\n")
    const built = contractOf(TEMPLATE_TEXT, outline)
    expect(built.contract.unaccounted.map(unit => unit.label)).toEqual(["ANNEX B – CHARGING SCHEDULE"])
    expect(outlineCoverageIssues(built.directory, built.documents).length).toBeGreaterThan(0)
  })

  test("a document with no Template line carries no part of another document's template", () => {
    const outline = [TITLE, ...[PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration, PARTS.annexA].map(heading),
      "# Covering letter", "## Charging schedule", "Explains the charging schedule."].join("\n")
    const built = contractOf(TEMPLATE_TEXT, outline)
    expect(built.contract.unaccounted.map(unit => unit.label)).toEqual(["ANNEX B – CHARGING SCHEDULE"])
  })

  test("a part another document drafted from the same template carries is not expected here", () => {
    const outline = [TITLE, ...[PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration].map(heading),
      "# Help desk annexes", `Template: template.docx`, heading(PARTS.annexA), heading(PARTS.annexB)].join("\n")
    const built = contractOf(TEMPLATE_TEXT, outline)
    expect(built.contract.places.get(5)).toEqual({ kind: "document", title: "Help desk annexes" })
    expect(templateContract(built.directory, built.documents, 1)!.places.get(1)).toEqual({ kind: "document", title: "Help desk agreement" })
    expect(outlineCoverageIssues(built.directory, built.documents)).toEqual([])
    expect(gapsOf(built, PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration)).toEqual([])
  })
})

// A rendered draft whose part labels are not printed: the title, then each
// part's text, its own headings included, attributed to its part.
const unlabelled = (texts: string[]): ParagraphRow[] => [{ pid: "p-title", sha256: "", text: "Help desk agreement", draft_part: null, draft_heading: true },
  ...texts.flatMap((text, part) => textParagraphs(text).map((paragraph, at) => ({ pid: `p-${part}-${at}`, sha256: "", text: paragraph, draft_part: part, draft_heading: false })))]
// Each part's text under its own heading, the template's, where its label was.
const OWN_HEADINGS = Object.values(PARTS).map(text => text.replace(/^## /, "### "))
const RENAMED = OWN_HEADINGS.map(text => text.replace("### 2. Charges", "### Fees"))

describe("template coverage", () => {
  test("in a revision, a part of the original under an unprinted label with no heading naming it lacks its heading; a printed heading or label holds it, and a new draft adapts headings freely", () => {
    const built = contractOf(TEMPLATE_TEXT, OUTLINE)
    const revision = { ...built.documents[0]!, mode: "revise" as const }
    const gaps = (texts: string[], outline: Parameters<typeof rowsBody>[0] = revision) => contractGaps(built.contract, rowsBody(outline, unlabelled(texts))).map(gap => [gap.unit.label, gap.gap])
    expect(gaps(OWN_HEADINGS)).toEqual([])
    expect(gaps(RENAMED)).toEqual([["2 Charges", "unheaded"]])
    expect(gaps(OWN_HEADINGS.map(text => text.replace("### 2. Charges\n\n", "")))).toEqual([["2 Charges", "unheaded"]])
    expect(gaps(RENAMED, { ...revision, parts: revision.parts.map((part, at) => at === 1 ? { ...part, display_heading: "Fees" } : part) })).toEqual([])
    expect(gaps(RENAMED, { ...revision, parts: revision.parts.map(({ display_heading: _, ...part }) => part) })).toEqual([])
    expect(gaps(RENAMED, built.documents[0]!)).toEqual([])
  })

  test("a revision's part without its heading is a draft open item; in the final text it is an omission, or a removal where an applied amendment took the heading out", () => {
    const built = contractOf(TEMPLATE_TEXT, OUTLINE)
    writeLeanPlan(built.directory, built.documents.map(document => ({ ...document, mode: "revise" as const })))
    const item = (rows: OpenItemRows) => rows.map(row => [row.kind, row.detail])
    type OpenItemRows = ReturnType<typeof draftTemplateItems>
    expect(item(draftTemplateItems(built.directory, "d1", unlabelled(RENAMED))))
      .toEqual([["omission", 'The template part "2 Charges" of "template.docx" is in the draft without its heading.']])
    const draft = unlabelled(OWN_HEADINGS)
    const heading = draft.find(row => row.text === "2. Charges")!
    const final = draft.map(row => row.pid === heading.pid ? { ...row, text: "Fees" } : row)
    expect(item(templateItems(built.directory, "d1", draft, final, [], new Set())))
      .toEqual([["omission", 'The heading of the template part "2 Charges" of "template.docx" is not in the final text; check whether it should be restored.']])
    const renamed = { id: "a1", pid: heading.pid, existing_sha256: "", updated: "Fees", reason: "The parties call the charges fees.", sources: [] } as unknown as Amendment
    expect(item(templateItems(built.directory, "d1", draft, final, [renamed], new Set(["a1"])))).toEqual([["template_removal",
      'The heading of the template part "2 Charges" of "template.docx" is taken out by amendment a1. Reason given: a1: The parties call the charges fees.']])
  })

  test("a part is measured in the part the contract gives it, under any heading and in any position", () => {
    const built = contractOf(TEMPLATE_TEXT, OUTLINE)
    expect(gapsOf(built, ...Object.values(PARTS))).toEqual([])
    expect(gapsOf(built, PARTS.annexB, PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration, PARTS.annexA)).toEqual([])
    expect(gapsOf(built, PARTS.scope, PARTS.charges, PARTS.duration, PARTS.annexA, PARTS.annexB)).toEqual([["3 Service Credits", "missing"]])
    const renamed = contractOf(TEMPLATE_TEXT, [TITLE, heading(PARTS.scope), "## Fees", "Carries the template's charges clause.",
      heading(PARTS.credits), heading(PARTS.duration), heading(PARTS.annexA), heading(PARTS.annexB)].join("\n"))
    expect(gapsOf(renamed, PARTS.scope, PARTS.charges.replace("## 2. Charges", "## Fees"), PARTS.credits, PARTS.duration, PARTS.annexA, PARTS.annexB)).toEqual([])
  })

  test("a placeholder or an open-item marker carries none of a part's text", () => {
    const built = contractOf(TEMPLATE_TEXT, OUTLINE)
    const marked = `${heading(PARTS.credits)}\n\n[open item: The service credits clause is left out because no response times are agreed; confirm before signature.]`
    const placeholder = `${heading(PARTS.annexB)}\n\n[insert the charging schedule agreed between the parties]`
    expect(gapsOf(built, PARTS.scope, PARTS.charges, marked, PARTS.duration, PARTS.annexA, placeholder))
      .toEqual([["3 Service Credits", "missing"], ["ANNEX B – CHARGING SCHEDULE", "missing"]])
  })

  test("a sentence naming a part in passing does not carry it", () => {
    const built = contractOf(TEMPLATE_TEXT, OUTLINE)
    const mention = `${PARTS.charges}\n\nThe charging schedule in Annex B applies.`
    expect(gapsOf(built, PARTS.scope, mention, PARTS.credits, PARTS.duration, PARTS.annexA)).toEqual([["ANNEX B – CHARGING SCHEDULE", "missing"]])
  })

  test("a schedule headed with the name of a clause is told apart from that clause", () => {
    const template = [
      "SUPPORT AGREEMENT",
      "1. Services\nThe supplier shall provide the support services described in Schedule 1 during the term, with reasonable skill and care and in line with good industry practice.",
      "2. Payment\nThe customer shall pay each invoice within thirty days of receipt, and late amounts carry simple interest at the agreed rate until they are paid in full.",
      "SCHEDULE 1 – SERVICES\nThe supplier shall answer calls on working days, visit the site within two working days of a reported fault, keep spare parts for each covered machine, service each machine twice a year, and keep a written record of every visit and repair for the customer to inspect on request.",
    ].join("\n\n")
    expect(parseTemplateUnits(template).find(unit => unit.label === "SCHEDULE 1 – SERVICES")!.alternatives.map(words => words.join(" "))).toEqual(["schedule 1 service", "schedule 1"])
    const services = "## 1. Services\n\nThe supplier provides the support services described in Schedule 1 throughout the term, with reasonable skill and care."
    const payment = "## 2. Payment\n\nThe customer pays each invoice within thirty days, with interest on late amounts."
    const schedule = "## Schedule 1 – Services\n\nThe supplier answers calls on working days, visits within two working days of a fault, keeps spare parts, services each machine twice a year and records every visit."
    const built = contractOf(template, ["# Support agreement", heading(services), heading(payment), heading(schedule)].join("\n"))
    expect(gapsOf(built, services, payment)).toEqual([["SCHEDULE 1 – SERVICES", "missing"]])
    expect(gapsOf(built, services, payment, schedule)).toEqual([])
    expect(gapsOf(built, payment, schedule)).toEqual([["1 Services", "missing"]])
  })

  test("the rendered document's title heads no part, even when it holds the name of one", () => {
    const outline = ["# Goods Delivery Agreement", "## Background", "## 1. Orders", "## 2. Delivery", "## 3. Price", "## 4. Invoicing"].join("\n")
    const built = contractOf(DIVIDED, outline)
    const sections = DIVIDED.split("\n\n").filter(block => block.includes("\n")).map(block => block.split("\n"))
    const rows = (blocks: string[][]) => ["Goods Delivery Agreement", ...blocks.flat()].map((text, at) => ({ pid: `p${at}`, sha256: "", text }))
    const gaps = (draft: ReturnType<typeof rows>) => contractGaps(built.contract, rowsBody(built.documents[0]!, draft)).map(gap => [gap.unit.label, gap.gap])
    expect(gaps(rows(sections))).toEqual([])
    expect(gaps(rows(sections.filter(block => block[0] !== "2. Delivery")))).toEqual([["2 Delivery", "missing"]])
  })

  test("a part drafted in a small share of its template words is thin, also where one outlined part carries two template parts", () => {
    const built = contractOf(TEMPLATE_TEXT, OUTLINE)
    expect(gapsOf(built, PARTS.scope, PARTS.charges, "## 3. Service Credits\n\nCredits apply as agreed.", PARTS.duration, PARTS.annexA, PARTS.annexB))
      .toEqual([["3 Service Credits", "thin"]])
    const hire = ["EQUIPMENT HIRE AGREEMENT",
      "1. Equipment\nThe owner hires to the hirer the items of equipment described in the order, each in working condition and fit for its normal purpose, together with any manuals, keys and accessories supplied with it. The hirer takes the equipment as described and may not substitute other items without consent.",
      "2. Hire Charges\nThe hirer pays the weekly hire charge set out in the order, invoiced in advance at the start of each week, with any part week charged pro rata at the daily rate. Charges exclude sales taxes, which the hirer pays in addition at the rate in force, and are payable in full without set off.",
      "3. Late Payment\nIf the hirer does not pay an invoice when it falls due, the owner may charge interest on the unpaid amount at the agreed rate from the due date until payment, may suspend further deliveries until the account is paid, and may recover its reasonable costs of collecting the debt.",
    ]
    const combined = contractOf(hire.join("\n\n"), ["# Equipment hire agreement", "## 1. Equipment", "## 2. Hire Charges and Late Payment"].join("\n"), "hire.docx")
    expect(gapsOf(combined, `## 1. Equipment\n\n${hire[1]!.split("\n")[1]}`, "## 2. Hire Charges and Late Payment\n\nThe hirer pays the weekly charge. Unpaid invoices carry interest at the agreed rate."))
      .toEqual([["2 Hire Charges", "thin"], ["3 Late Payment", "thin"]])
  })
})
