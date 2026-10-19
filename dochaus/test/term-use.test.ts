import { describe, expect, test } from "bun:test"
import { findTermUse } from "../lib/search/term-use"

describe("findTermUse", () => {
  test("returns the first whole-word occurrence after skipping embedded ones", () => {
    const text = "Terms, Term_1, Term2 and Termination precede the Term."
    expect(findTermUse(text, "Term")).toBe(text.lastIndexOf("Term"))
  })

  test("treats punctuation, including apostrophes, as a boundary", () => {
    expect(findTermUse("Each Party's address", "Party")).toBe(5)
    expect(findTermUse("Each Party’s address", "Party")).toBe(5)
    expect(findTermUse("a Non-Party lender", "Party")).toBe(6)
  })

  test("is case-sensitive and treats non-ASCII letters as word characters", () => {
    expect(findTermUse("any third party", "Party")).toBe(-1)
    expect(findTermUse("PartyÉ", "Party")).toBe(-1)
  })

  test("never matches an empty term", () => {
    expect(findTermUse("any text", "")).toBe(-1)
  })
})
