// Run from dochaus/: bun test lib/markdown.test.ts
import { describe, expect, test } from "bun:test"
import { lintRenderedDraftMarkdown, renderMarkdownToDocx } from "./render"
import { strFromU8, unzipSync } from "fflate"
import {
  inspectDraftMarkdownNamespace,
  lintDraftMarkdown,
  lintTemplateBody,
  splitBlocks,
} from "./markdown"

const CLEAN = `# MUTUAL NDA

Between [insert first party] and [insert second party].

## 1. Purpose

The Parties wish to explore [insert business opportunity].

## 2. Term

[insert term in years] years.

## 3. Non-Solicitation [optional: non-solicitation]

For [insert non-solicit period in months] months.

## 4. Entire Agreement

Whole deal.`

describe("splitBlocks", () => {
  test("every line becomes its own block", () => {
    expect(splitBlocks("## 3. Compensation\n(a) Base Salary. Text.\n(b) Bonus. Text.")).toBe(
      "## 3. Compensation\n\n(a) Base Salary. Text.\n\n(b) Bonus. Text.",
    )
  })

  test("existing blank lines collapse to one boundary", () => {
    expect(splitBlocks("# Title\n\n\nBody.")).toBe("# Title\n\nBody.")
  })

  test("html centering wrappers do not render as literal document text", () => {
    expect(splitBlocks("## Instrument Title\n<center>\n\n**CONSENT TO ADOPTION**\n\n</center>\n## 1. Identification")).toBe(
      "## Instrument Title\n\n**CONSENT TO ADOPTION**\n\n## 1. Identification",
    )
  })
})

const STYLE = { dateFormat: "month-day-year", numberStyle: "words-and-numerals" } as const

const DRAFT = `# EMPLOYMENT AGREEMENT

This Employment Agreement (this "Agreement") is entered into as of June 11, 2026 by Acme Corp (the "Company") and Jane Roe (the "Employee").

## 1. Term

The Employee's employment begins on the date above and continues for two (2) years unless terminated under Section 2.

## 2. Termination

The Company or the Employee may terminate this Agreement on thirty (30) days written notice. Nothing limits rights under California Labor Code Section 2870.`

describe("lintDraftMarkdown", () => {
  test("clean draft passes", () => {
    expect(lintDraftMarkdown(DRAFT, STYLE)).toEqual([])
  })

  test("recognizes a capitalized term defined with paired single quotes", () => {
    const draft = DRAFT.replace(
      'Acme Corp (the "Company")',
      "Acme Corp. As used in this Agreement, the term 'Company' means Acme Corp",
    )

    expect(lintDraftMarkdown(draft, STYLE)).toEqual([])
  })

  test("defined but unused term flagged", () => {
    const problems = lintDraftMarkdown(DRAFT.replace('(the "Company")', '(the "Company", guaranteed by the "Guarantor")'), STYLE)
    expect(problems).toEqual(['"Guarantor" is defined but never used'])
  })

  test("repeated undefined term flagged, single mention not", () => {
    const problems = lintDraftMarkdown(
      DRAFT + "\n\nAwards vest under the Equity Plan. The Equity Plan is discretionary. The Watermark appears once.",
      STYLE,
    )
    expect(problems).toEqual(['"Equity Plan" reads like a defined term but is never defined'])
  })

  test("broken cross-reference flagged, statute citation not", () => {
    const problems = lintDraftMarkdown(DRAFT.replace("under Section 2", "under Sections 2 and 9"), STYLE)
    expect(problems).toEqual(["cross-reference to Section 9, but no heading is numbered 9"])
  })

  test("bare numeral duration flagged under words-and-numerals", () => {
    const problems = lintDraftMarkdown(DRAFT.replace("thirty (30) days", "30 days"), STYLE)
    expect(problems.some((p) => p.includes('"30 days"'))).toBe(true)
  })

  test("word-numeral pairs flagged under numerals", () => {
    const problems = lintDraftMarkdown(DRAFT, { ...STYLE, numberStyle: "numerals" })
    expect(problems.some((p) => p.includes('"two (2)"') && p.includes('"thirty (30)"'))).toBe(true)
  })

  test("wrong date format flagged", () => {
    const problems = lintDraftMarkdown(DRAFT, { ...STYLE, dateFormat: "iso" })
    expect(problems.some((p) => p.includes('"June 11, 2026"'))).toBe(true)
  })

  test("bare back-reference to a statutory section number not flagged", () => {
    expect(lintDraftMarkdown(DRAFT + "\n\nEmployee acknowledges notification of the provisions of Section 2870.", STYLE)).toEqual([])
  })

  test("statutory section chains and full law names are not internal cross-references", () => {
    const draft = DRAFT + `

The board acts under DGCL Section 102 and Section 151.

The incorporator acts under Sections 107 and 108 of the Delaware General Corporation Law.

Section 83(b) Advisory Follow-Up.

A tax election may be filed under Section 83(b) of the Internal Revenue Code.`

    expect(lintDraftMarkdown(draft, STYLE)).toEqual([])
  })

  test("documents without the defined-term convention skip term checks", () => {
    const memo = `# COVER MEMO\n\nI prepared the draft from the Company's standard template. Please confirm the Company's signature block and the Equity Plan treatment. The Equity Plan vests annually on June 11, 2026.`
    expect(lintDraftMarkdown(memo, STYLE)).toEqual([])
  })

  test("bare blank placeholder flagged", () => {
    const problems = lintDraftMarkdown(DRAFT.replace("Jane Roe", "[___]"), STYLE)
    expect(problems.some((p) => p.includes("[___]"))).toBe(true)
  })
})

describe("lintRenderedDraftMarkdown", () => {
  test("lints a Word projection's visible text line for line", () => {
    const projected = [
      "## 1. Services",
      "",
      'The "Alpha Marker", "Alpha Signal", and "Alpha Token" are defined here. The appointed supplier \\(Lead Vendor\\) acts for both parties.',
      "The Alpha Marker, the Alpha Signal and the Alpha Token apply. The Lead Vendor reports monthly. The Lead Vendor keeps records.",
      "Notice follows Section 2\\(a\\) of the Act. The fee is \\[\\_\\_\\_\\].",
      "",
      "\\#4 The Missing Marker applies. The Missing Marker continues.",
    ].join("\n")

    expect(lintRenderedDraftMarkdown(projected)).toEqual([
      '"Missing Marker" reads like a defined term but is never defined',
      "bare [___] left in the document — fill it or use a descriptive [insert ...] placeholder",
    ])
  })
})

// A Word projection escapes punctuation that has no visible form. The scanner
// must diagnose the projected body exactly as it diagnoses the same source.
describe("Word-projected namespace text", () => {
  const DEFINED =
    'The "Alpha Marker", "Alpha Signal", and "Alpha Token" are defined here. The Alpha Marker, the Alpha Signal and the Alpha Token apply.'

  test.each([
    [
      "a hyphenated quoted definition",
      `${DEFINED} The "Rate-Card" is defined here. The Rate-Card applies. The Rate-Card is fixed.`,
      `${DEFINED} The "Rate\\-Card" is defined here. The Rate\\-Card applies. The Rate\\-Card is fixed.`,
    ],
    [
      "a parenthesized definition",
      `${DEFINED} The supplier (Lead Vendor) acts. The Lead Vendor reports. The Lead Vendor keeps records.`,
      `${DEFINED} The supplier \\(Lead Vendor\\) acts. The Lead Vendor reports. The Lead Vendor keeps records.`,
    ],
    [
      "an external subsection reference",
      `## 1. Alpha\n\n${DEFINED} Notice follows Section 2(a) of the Act.`,
      `## 1. Alpha\n\n${DEFINED} Notice follows Section 2\\(a\\) of the Act.`,
    ],
    [
      "a literal code span",
      `${DEFINED} The flag \`\\-\` stays literal.`,
      `${DEFINED} The flag \`\\-\` stays literal.`,
    ],
  ])("reads %s on its visible text", (_name, source, projected) => {
    expect(inspectDraftMarkdownNamespace(source)).toEqual([])
    expect(inspectDraftMarkdownNamespace(projected)).toEqual([])
  })

  test("activates term checks from escaped parenthesized definitions as the source does", () => {
    const source = [
      "The parties are the seller (Alpha), the buyer (Beta) and the agent (Gamma). Alpha, Beta and Gamma sign.",
      "The Unknown Label applies. The Unknown Label continues.",
    ].join("\n")
    const projected = source.replace(/\((Alpha|Beta|Gamma)\)/g, "\\($1\\)")
    expect(projected).not.toBe(source)
    const expected = [{ code: "undefined_term", line: 2, message: '"Unknown Label" reads like a defined term but is never defined' }]
    expect(inspectDraftMarkdownNamespace(source)).toEqual(expected)
    expect(inspectDraftMarkdownNamespace(projected)).toEqual(expected)
  })

  test("keeps a paragraph that starts with an escaped hash a paragraph", () => {
    const projected = `${DEFINED}\n\n\\#4 The Missing Marker applies. The Missing Marker continues.`
    expect(inspectDraftMarkdownNamespace(projected)).toEqual([
      { code: "undefined_term", line: 3, message: '"Missing Marker" reads like a defined term but is never defined' },
    ])
  })
})

describe("inspectDraftMarkdownNamespace", () => {
  test("returns only typed term and cross-reference diagnostics", () => {
    const issues = inspectDraftMarkdownNamespace(
      DRAFT
        .replace("under Section 2", "under Section 9")
        .replace(
          '(the "Company")',
          '(the "Company", guaranteed by the "Guarantor")',
        ),
    )

    expect(issues).toContainEqual({
      code: "unused_term",
      line: 3,
      message: '"Guarantor" is defined but never used',
    })
    expect(issues).toContainEqual({
      code: "broken_cross_reference",
      line: 7,
      message: "cross-reference to Section 9, but no heading is numbered 9",
    })
    expect(issues.some((issue) => issue.message.includes("date"))).toBe(false)
  })

  test("reports the exact first diagnostic line for repeated undefined terms", () => {
    const issues = inspectDraftMarkdownNamespace(
      DRAFT +
        "\n\nThe Equity Plan applies.\n\nA later provision also invokes the Equity Plan.",
    )

    expect(issues).toContainEqual({
      code: "undefined_term",
      line: 13,
      message: '"Equity Plan" reads like a defined term but is never defined',
    })
  })

  test("uses the document title when classifying apparent defined terms", () => {
    const issues = inspectDraftMarkdownNamespace(`
The "Alpha Marker", "Alpha Signal", and "Alpha Token" are defined here.
The Agreement applies. This Agreement continues.
The Missing Marker applies. The Missing Marker continues.
`, { documentTitle: "Settlement Agreement" })

    expect(issues.filter((issue) => issue.code === "undefined_term")).toEqual([
      {
        code: "undefined_term",
        line: 4,
        message: '"Missing Marker" reads like a defined term but is never defined',
      },
    ])
  })

  test("does not treat dates, acronyms, agencies, or compounds ending in a defined term as undefined", () => {
    const issues = inspectDraftMarkdownNamespace(
      DRAFT + `

The January deadline follows an earlier January filing.

A SAFE financing may close. The SAFE financing remains subject to approval.

The Internal Revenue Service issued guidance. The Internal Revenue Service may update it.

The Bylaws of the Company are attached. The Bylaws of the Company control.`
    )

    expect(issues.filter((issue) => issue.code === "undefined_term")).toEqual([])
  })

  test("still reports a repeated inconsistent party label", () => {
    const issues = inspectDraftMarkdownNamespace(
      DRAFT + "\n\nThe Issuer may act. The Issuer remains responsible.",
    )

    expect(issues).toContainEqual({
      code: "undefined_term",
      line: 13,
      message: '"Issuer" reads like a defined term but is never defined',
    })
  })

  test("recognizes unquoted parenthetical aliases and section aliases in headings", () => {
    const issues = inspectDraftMarkdownNamespace(`
## 1. Parties

Acme Holdings, Inc. (Acme) and Jamie Roe (Employee) enter this instrument. Acme shall pay Employee. Acme and Employee remain bound.

The "Payment Date", "Closing Date", and "Review Period" apply under Section 16.

## 6. Sections 5 and 16 — Review Mechanics

The Payment Date follows the Closing Date after the Review Period.
`)

    expect(issues).toEqual([])
  })

  test("accepts declared external section identities without weakening local checks", () => {
    const memo = "## 1. Advice\n\nReview Section 010 and Section 4 before signing."
    expect(inspectDraftMarkdownNamespace(memo)).toEqual([
      {
        code: "broken_cross_reference",
        line: 3,
        message: "cross-reference to Section 010, but no heading is numbered 010",
      },
      {
        code: "broken_cross_reference",
        line: 3,
        message: "cross-reference to Section 4, but no heading is numbered 4",
      },
    ])
    expect(inspectDraftMarkdownNamespace(memo, {
      externalSectionNumbers: ["010", "4"],
    })).toEqual([])
  })

  test("recognizes a closed bold section label and a Markdown heading as local targets", () => {
    const issues = inspectDraftMarkdownNamespace(`**Section 4 — Scope**

The scope is described here.

## Section 5. Payment

Payment follows Section 4 and Section 5, subject to Section 9.`)

    expect(issues).toEqual([{
      code: "broken_cross_reference",
      line: 7,
      message: "cross-reference to Section 9, but no heading is numbered 9",
    }])
  })

  test("recognizes a titled bold heading span before prose on the same line", () => {
    const issues = inspectDraftMarkdownNamespace(`**Section 4 — Scope.** The parties agree to the stated scope.

**Section 5 Payment Terms — Timing.** Payment follows the agreed schedule.

Review Section 4 and Section 5, then correct Section 9.`)

    expect(issues).toEqual([{
      code: "broken_cross_reference",
      line: 5,
      message: "cross-reference to Section 9, but no heading is numbered 9",
    }])
  })

  test("does not promote ambiguous plain or partly bold section prose into targets", () => {
    const issues = inspectDraftMarkdownNamespace(`## 1. Scope

Section 9: see the referenced material before execution.

**Section 10**. This cross-reference has not been checked.

Section 11. Payment terms are stated elsewhere.

**Section 12.** This is punctuation after a number, not a titled heading.`)

    expect(issues).toEqual([
      {
        code: "broken_cross_reference",
        line: 3,
        message: "cross-reference to Section 9, but no heading is numbered 9",
      },
      {
        code: "broken_cross_reference",
        line: 5,
        message: "cross-reference to Section 10, but no heading is numbered 10",
      },
      {
        code: "broken_cross_reference",
        line: 7,
        message: "cross-reference to Section 11, but no heading is numbered 11",
      },
      {
        code: "broken_cross_reference",
        line: 9,
        message: "cross-reference to Section 12, but no heading is numbered 12",
      },
    ])
  })

  test("recognizes lowercase section in external citations and local references", () => {
    const issues = inspectDraftMarkdownNamespace(`## 1. Scope

Example Code section 44 applies. The same rule is stated in section 44 of the Example Code.

Review section 1 and section 9 before signing.`)

    expect(issues).toEqual([{
      code: "broken_cross_reference",
      line: 5,
      message: "cross-reference to Section 9, but no heading is numbered 9",
    }])
  })
})

describe("lintTemplateBody", () => {
  test("clean body passes", () => {
    expect(lintTemplateBody(CLEAN)).toEqual([])
  })

  test("duplicate placeholder flagged", () => {
    const problems = lintTemplateBody(CLEAN.replace("[insert second party]", "[insert first party]"))
    expect(problems.some((p) => p.includes("[insert first party]") && p.includes("more than once"))).toBe(true)
  })

  test("bare blank flagged", () => {
    const problems = lintTemplateBody(CLEAN.replace("[insert term in years]", "[___]"))
    expect(problems.some((p) => p.includes("[___]"))).toBe(true)
  })

  test("heading numbering gap flagged", () => {
    const problems = lintTemplateBody(CLEAN.replace("## 4. Entire Agreement", "## 5. Entire Agreement"))
    expect(problems).toEqual(['heading numbering breaks at "5." — expected 4'])
  })

  test("duplicate optional marker flagged", () => {
    const problems = lintTemplateBody(
      CLEAN.replace("## 4. Entire Agreement", "## 4. No Hire [optional: non-solicitation]"),
    )
    expect(problems.some((p) => p.includes("[optional: non-solicitation]") && p.includes("more than once"))).toBe(true)
  })
})
