import { afterEach, describe, expect, test } from "bun:test"
import { writeLeanPlan } from "./plan"
import { outlineCoverageIssues, templateContract } from "./template-contract"
import { contractGaps, draftBody, templateItems, textParagraphs } from "./template-coverage"
import { PARTS, TEMPLATE_TEXT, TITLE, contractOf, heading } from "./template-test-support"
import { cleanupMatters } from "./test-support"

// Placement across the planned documents, and what a Left out line records,
// on invented templates.

afterEach(cleanupMatters)

const ANNEX_B = 6
const MAIN = [TITLE, heading(PARTS.scope), heading(PARTS.charges), "Pays the charges in the charging schedule.", heading(PARTS.credits), heading(PARTS.duration), heading(PARTS.annexA)]

describe("placement across planned documents", () => {
  test("a part goes to another document only when that document's own outline holds it in its parts", () => {
    // The other document heads the part: it holds it there, so this one need not.
    const headed = contractOf(TEMPLATE_TEXT, [...MAIN, "# Charging schedule", "Template: template.docx", heading(PARTS.annexB), "## Extra hours and expenses"].join("\n"))
    expect(headed.contract.places.get(ANNEX_B)).toEqual({ kind: "document", title: "Charging schedule" })
    expect(templateContract(headed.directory, headed.documents, 1)!.places.get(ANNEX_B)).toEqual({ kind: "parts", parts: [0] })
    expect(outlineCoverageIssues(headed.directory, headed.documents)).toEqual([])

    // The other document is only titled after the part: this one's scope line holds it, and the other points here.
    const titled = contractOf(TEMPLATE_TEXT, [...MAIN, "# Charging schedule", "Template: template.docx", "## Monthly fee", "## Extra hours and expenses"].join("\n"))
    expect(titled.contract.places.get(ANNEX_B)).toEqual({ kind: "parts", parts: [1] })
    expect(templateContract(titled.directory, titled.documents, 1)!.places.get(ANNEX_B)).toEqual({ kind: "document", title: "Help desk agreement" })

    // The other document heads the part but also leaves it out, so its own contract holds it nowhere:
    // this one must account for it itself.
    const left = contractOf(TEMPLATE_TEXT, [TITLE, ...[PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration, PARTS.annexA].map(heading),
      "# Help desk annexes", "Template: template.docx", 'Left out: "Annex B" – the charges are fixed in the agreement.', heading(PARTS.annexA), heading(PARTS.annexB)].join("\n"))
    expect(left.contract.unaccounted.map(unit => unit.label)).toEqual(["ANNEX B – CHARGING SCHEDULE"])
    expect(left.contract.places.get(ANNEX_B)).toEqual({ kind: "parts", parts: [4], assumed: true })
    expect(templateContract(left.directory, left.documents, 1)!.places.get(ANNEX_B)?.kind).toBe("left_out")
  })

  test("a document's title places no part, so a part its headings leave out stays outline feedback and is measured", () => {
    const clauses = [
      "1. Consultancy\nThe consultant shall provide the advice described in the brief, attend the meetings the client reasonably asks for, and deliver a written report at the end of each stage, using the care and skill expected of a competent adviser in the field.",
      "2. Charges\nThe client shall pay the day rate set out in the brief for each day worked, invoiced monthly in arrears, together with travel costs approved in advance. Invoices are payable within thirty days and exclude sales taxes, which the client pays in addition.",
      "3. Confidentiality\nEach party keeps the other's confidential information secret, uses it only for the engagement, and returns or destroys it when the engagement ends, except where the law requires disclosure or the information is already public through no fault of the receiving party.",
    ]
    const headings = ["## 1. Advice and reports", "## 2. Charges", "## 3. Confidentiality"]
    const draft = draftBody([{ part: 0, text: `${headings[0]}\n\nThe consultant reports as agreed.` },
      ...[1, 2].map(part => ({ part, text: `${headings[part]}\n\n${clauses[part]!.split("\n")[1]}` }))])
    // A title rewording the template's title, and one matching a title the template's cover line hides.
    for (const [template, title] of [[["CONSULTANCY AGREEMENT", ...clauses], "Consultancy contract"], [["Draft for discussion", "CONSULTANCY AGREEMENT", ...clauses], "Consultancy agreement"]] as const) {
      const built = contractOf(template.join("\n\n"), [`# ${title}`, ...headings].join("\n"), "consultancy.docx")
      expect(built.contract.unaccounted.map(unit => unit.label)).toContain("1 Consultancy")
      expect(outlineCoverageIssues(built.directory, built.documents)).toHaveLength(1)
      expect(contractGaps(built.contract, draft).map(gap => [gap.unit.label, gap.gap])).toContainEqual(["1 Consultancy", "thin"])
    }
  })
})

describe("left out parts", () => {
  test("a part both left out and headed in the plan is an outline issue", () => {
    const own = contractOf(TEMPLATE_TEXT, [TITLE, 'Left out: "3. Service Credits" – no response times are agreed.', ...Object.values(PARTS).map(heading)].join("\n"))
    expect(outlineCoverageIssues(own.directory, own.documents)).toEqual([
      'The "Left out:" line "3. Service Credits" of "Help desk agreement" leaves out the template part "3 Service Credits", which "3. Service Credits" also names. ' +
      'Keep the "Left out:" line only if no part carries it; otherwise drop that line.',
    ])
    const sibling = contractOf(TEMPLATE_TEXT, [TITLE, 'Left out: "Annex B" – drafted on its own.', ...[PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration, PARTS.annexA].map(heading),
      "# Charging schedule", "Template: template.docx", heading(PARTS.annexB)].join("\n"))
    expect(outlineCoverageIssues(sibling.directory, sibling.documents)).toEqual([
      'The "Left out:" line "Annex B" of "Help desk agreement" leaves out the template part "ANNEX B – CHARGING SCHEDULE", which "Charging schedule" also names. ' +
      'Keep the "Left out:" line only if no part carries it; otherwise drop that line.',
    ])
  })

  test("a part left out with no reason is an omission, not a recorded removal", () => {
    const itemsFor = (line: string) => {
      const built = contractOf(TEMPLATE_TEXT, [TITLE, line, ...[PARTS.scope, PARTS.charges, PARTS.duration, PARTS.annexA, PARTS.annexB].map(heading)].join("\n"))
      writeLeanPlan(built.directory, built.documents)
      const rows = textParagraphs([TITLE, PARTS.scope, PARTS.charges, PARTS.duration, PARTS.annexA, PARTS.annexB].join("\n\n"))
        .map((text, at) => ({ pid: `p${at}`, sha256: "", text }))
      return templateItems(built.directory, "d1", rows, rows, [], new Set()).map(item => [item.kind, item.detail])
    }
    expect(itemsFor('Left out: "3. Service Credits"')).toEqual([
      ["omission", 'The template part "3 Service Credits" of "template.docx" is left out by the outline, which gives no reason.'],
    ])
    expect(itemsFor('Left out: "3. Service Credits" – no response times are agreed.').map(item => item[0])).toEqual(["template_removal"])
  })
})
