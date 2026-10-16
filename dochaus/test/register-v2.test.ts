import { expect, test } from "bun:test"
import { normalizeToolMetadata } from "../tool/register-v2"

test("native legal tool metadata preserves citation data and omits absent optional fields", () => {
  const input = { citations: [{ document: "source.docx", section: "1", structuralPath: undefined, party: undefined, score: 0.5, superseded_by: undefined }], pending: [], title: "Sources", enabled: false, count: 0, absent: null }
  expect(normalizeToolMetadata(input)).toEqual({ citations: [{ document: "source.docx", section: "1", score: 0.5 }], pending: [], title: "Sources", enabled: false, count: 0, absent: null })
  expect(input.citations[0]).toHaveProperty("structuralPath")
})

test("invalid tool metadata cannot become a successful empty result", () => {
  const circular: Record<string, unknown> = {}
  circular.self = circular
  expect(() => normalizeToolMetadata(circular)).toThrow()
  expect(() => normalizeToolMetadata({ unsupported: 1n })).toThrow()
})
