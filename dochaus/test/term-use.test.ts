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

  test("treats a combining mark after the term as part of the word", () => {
    expect(findTermUse("Cafe\u0301 x", "Cafe")).toBe(-1)
    expect(findTermUse("x \u0301Cafe", "Cafe")).toBe(-1)
  })

  test("reads a whole surrogate pair on either side of the term", () => {
    expect(findTermUse("\u{1D400}Party", "Party")).toBe(-1)
    expect(findTermUse("Party\u{1D400}", "Party")).toBe(-1)
    // A non-letter astral character is still a boundary.
    expect(findTermUse("\u{1F600}Party", "Party")).toBe(2)
  })

  test("never matches an empty term", () => {
    expect(findTermUse("any text", "")).toBe(-1)
  })
})
