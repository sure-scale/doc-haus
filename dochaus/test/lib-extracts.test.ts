import { describe, expect, test } from "bun:test"
import { normalize, normalizeWithMap } from "../lib/quote-verify"
import { extractValues } from "../lib/coverage-primitives"

describe("quote-verify helpers", () => {
  test("normalize collapses whitespace and maps curly quotes", () => {
    expect(normalize("foo  “bar”\n\tbaz")).toBe('foo "bar" baz')
  })

  test("normalizeWithMap returns text and an offset map back to raw", () => {
    const raw = "foo  bar"
    const { text, map } = normalizeWithMap(raw)
    expect(text).toBe("foo bar")
    expect(map.length).toBe(text.length)
    expect(map[4]).toBe(raw.lastIndexOf("bar"))
  })
})

describe("coverage-primitives extractValues", () => {
  test("extracts dollar amounts and percentages", () => {
    expect(extractValues("$50,000 and 12.5% of $1,000,000.00")).toEqual(
      expect.arrayContaining(["$50,000", "12.5%", "$1,000,000.00"]),
    )
  })

  test("returns empty array for text with no values", () => {
    expect(extractValues("no values here")).toEqual([])
  })
})
