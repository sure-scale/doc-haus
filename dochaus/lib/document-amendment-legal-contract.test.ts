import { expect, test } from "bun:test"
import { decodeDocumentAmendmentLegal, documentAmendmentLegalJsonSchema } from "./document-amendment-legal-contract"

const amendment = { before_quote: "30 days", replacement: "45 days", target_context: "Delivery paragraph", reason: "Requested wording." }

test("the model contract and strict decoder require explicit amendments and open items", () => {
  expect(documentAmendmentLegalJsonSchema.required).toEqual(["amendments", "open_items"])
  expect(documentAmendmentLegalJsonSchema.additionalProperties).toBe(false)
  expect(decodeDocumentAmendmentLegal({ amendments: [amendment], open_items: [] })).toEqual({ amendments: [amendment], open_items: [] })
  for (const value of [
    { amendments: [amendment] },
    { amendments: [], open_items: null },
    { amendments: [], open_items: [], approved: true },
    { amendments: [{ ...amendment, offset: 0 }], open_items: [] },
    { amendments: Array.from({ length: 25 }, () => amendment), open_items: [] },
    { amendments: [], open_items: ["x".repeat(8193)] },
    { amendments: [], open_items: [], paragraph_insertions: null },
  ]) expect(() => decodeDocumentAmendmentLegal(value)).toThrow()
})

test("the structured transport retains supported structural shapes without supplying missing facts", () => {
  const value = { amendments: [], open_items: ["Confirm the requested date."], paragraph_formats: [{ kind: "unsupported_formatting", request: "Change the line spacing", reason: "Needs a supported operation." }] }
  expect(decodeDocumentAmendmentLegal(value)).toEqual(value)
  expect(() => decodeDocumentAmendmentLegal({ ...value, paragraph_formats: [{ kind: "unsupported_formatting", reason: "Missing requested work." }] })).toThrow()
})
