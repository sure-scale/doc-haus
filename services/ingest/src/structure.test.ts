import { expect, test } from "bun:test"
import {
  extractAmendments,
  extractDefinedTerms,
  extractParties,
  extractSectionOps,
  extractSectionRefs,
} from "./structure"

test("unmarked prose does not acquire family-specific structure", () => {
  const text = "Background\n\nThis document records the requested information in narrative form."

  expect(extractDefinedTerms(text)).toEqual([])
  expect(extractSectionRefs(text)).toEqual([])
  expect(extractParties(text)).toEqual([])
  expect(extractAmendments(text)).toEqual([])
  expect(extractSectionOps(text)).toEqual([])
})

test("restated pattern produces a restate op", () => {
  const text = "Section 4 of the Agreement is hereby amended and restated in its entirety to read as follows:"
  const ops = extractSectionOps(text)
  expect(ops).toHaveLength(1)
  expect(ops[0]).toMatchObject({ sectionLabel: "4", operation: "restate" })
  expect(ops[0]!.charStart).toBe(0)
  expect(ops[0]!.charEnd).toBe(text.length)
})

test("amended pattern (without 'and restated') produces an amend op", () => {
  const text = "Section 7.2(a) is hereby amended by deleting 'X' and inserting 'Y' in lieu thereof."
  const ops = extractSectionOps(text)
  expect(ops).toHaveLength(1)
  expect(ops[0]).toMatchObject({ sectionLabel: "7.2(a)", operation: "amend" })
})

test("deleted pattern produces a delete op", () => {
  const text = "Section 8 of the Agreement is hereby deleted in its entirety."
  const ops = extractSectionOps(text)
  expect(ops).toHaveLength(1)
  expect(ops[0]).toMatchObject({ sectionLabel: "8", operation: "delete" })
})

test("stricken is treated as delete", () => {
  const text = "Section 8 is hereby stricken in its entirety."
  const ops = extractSectionOps(text)
  expect(ops[0]).toMatchObject({ sectionLabel: "8", operation: "delete" })
})

test("new Section added produces an add op", () => {
  const text = "A new Section 9 is hereby added to read as follows:"
  const ops = extractSectionOps(text)
  expect(ops).toHaveLength(1)
  expect(ops[0]).toMatchObject({ sectionLabel: "9", operation: "add" })
  // "A new " is captured by the case-insensitive optional group, so the audit
  // span starts at the sentence-initial "A" — matches what a lawyer would
  // highlight as the operative sentence.
  expect(ops[0]!.charStart).toBe(0)
})

test("comma list expands to one op per label", () => {
  const text = "Sections 4, 5, and 7 are hereby deleted in their entireties."
  const ops = extractSectionOps(text)
  expect(ops.map((o) => o.sectionLabel).sort()).toEqual(["4", "5", "7"])
  expect(ops.every((o) => o.operation === "delete")).toBe(true)
})

test("'and' list expands to one op per label", () => {
  const text = "Sections 8.1 and 8.2 are hereby amended to read as follows:"
  const ops = extractSectionOps(text)
  expect(ops.map((o) => o.sectionLabel).sort()).toEqual(["8.1", "8.2"])
})

test("qualifier 'of the Original Credit Agreement' is stripped from label", () => {
  const text = "Section 4 of the Original Credit Agreement is hereby amended and restated."
  const ops = extractSectionOps(text)
  expect(ops[0]?.sectionLabel).toBe("4")
})

test("non-formulaic phrasing produces no row", () => {
  const text = "The second sentence of Section 4 is struck and replaced with the following:"
  const ops = extractSectionOps(text)
  expect(ops).toEqual([])
})

test("schedule-referenced restatement still records an amend op (resolver handles missing replacement)", () => {
  // The regex detects that §4 is amended; the *replacement text* lives in an
  // external Schedule the amender does not re-index. The resolver surfaces this
  // honestly via the "replacement text not separately indexed" fallback.
  const text = "Section 4 is hereby amended as set forth in Schedule A."
  const ops = extractSectionOps(text)
  expect(ops).toHaveLength(1)
  expect(ops[0]).toMatchObject({ sectionLabel: "4", operation: "amend" })
})

test("multiple ops across one document", () => {
  const text = [
    "Section 4 is hereby amended to read as follows: ...",
    "Section 7 is hereby deleted in its entirety.",
    "A new Section 9 is hereby added to read as follows: ...",
  ].join("\n\n")
  const ops = extractSectionOps(text)
  expect(ops).toHaveLength(3)
  expect(ops.map((o) => `${o.sectionLabel}:${o.operation}`).sort()).toEqual([
    "4:amend",
    "7:delete",
    "9:add",
  ])
})

test("roman numeral and single-letter labels are captured", () => {
  const text = "Article IV is hereby amended and restated to read as follows:"
  const ops = extractSectionOps(text)
  expect(ops[0]?.sectionLabel).toBe("IV")
})

test("offsets index into the input text", () => {
  const text = "Preamble text.\n\nSection 4 is hereby amended to read as follows: body."
  const ops = extractSectionOps(text)
  expect(ops[0]?.charStart).toBe(text.indexOf("Section 4"))
  expect(text.slice(ops[0]!.charStart, ops[0]!.charEnd)).toMatch(/^Section 4 is hereby amended/)
})
